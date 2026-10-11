// Package ldaptest runs a small in-process LDAP directory for tests: simple
// bind, LDAPS or StartTLS, subtree search with the filter shapes ldapx sends,
// and modify. It is not a general LDAP server.
package ldaptest

import (
	"crypto/ecdsa"
	"crypto/elliptic"
	"crypto/rand"
	"crypto/tls"
	"crypto/x509"
	"crypto/x509/pkix"
	"fmt"
	"math/big"
	"net"
	"regexp"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/jimlambrt/gldap"
)

// Server is a running directory.
type Server struct {
	Host   string
	Port   int
	Client *tls.Config // trusts the server's test certificate

	mu       sync.Mutex
	entries  map[string]map[string][]string // DN → attributes
	bindDN   string
	password string
}

// Start serves entries; tlsFromStart chooses LDAPS, otherwise the listener
// is plain and accepts StartTLS.
func Start(t *testing.T, bindDN, password string, tlsFromStart bool, entries map[string]map[string][]string) *Server {
	t.Helper()
	serverTLS, clientTLS := testTLS(t)
	l, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	port := l.Addr().(*net.TCPAddr).Port
	_ = l.Close()
	s := &Server{Host: "127.0.0.1", Port: port, Client: clientTLS,
		entries: map[string]map[string][]string{}, bindDN: bindDN, password: password}
	for dn, a := range entries {
		s.entries[dn] = a
	}
	srv, err := gldap.NewServer()
	if err != nil {
		t.Fatal(err)
	}
	mux, err := gldap.NewMux()
	if err != nil {
		t.Fatal(err)
	}
	_ = mux.Bind(s.bind)
	_ = mux.ExtendedOperation(func(w *gldap.ResponseWriter, r *gldap.Request) {
		res := r.NewExtendedResponse(gldap.WithResponseCode(gldap.ResultSuccess))
		res.SetResponseName(gldap.ExtendedOperationStartTLS)
		_ = w.Write(res)
		_ = r.StartTLS(serverTLS)
	}, gldap.ExtendedOperationStartTLS)
	_ = mux.Search(s.search)
	_ = mux.Modify(s.modify)
	_ = srv.Router(mux)
	var opts []gldap.Option
	if tlsFromStart {
		opts = append(opts, gldap.WithTLSConfig(serverTLS))
	}
	go func() { _ = srv.Run(net.JoinHostPort(s.Host, fmt.Sprint(s.Port)), opts...) }()
	t.Cleanup(func() { _ = srv.Stop() })
	deadline := time.Now().Add(5 * time.Second)
	for !srv.Ready() && time.Now().Before(deadline) {
		time.Sleep(5 * time.Millisecond)
	}
	return s
}

// Attr returns an entry's attribute values (nil if absent).
func (s *Server) Attr(dn, name string) []string {
	s.mu.Lock()
	defer s.mu.Unlock()
	for k, v := range s.entries[dn] {
		if strings.EqualFold(k, name) {
			return append([]string(nil), v...)
		}
	}
	return nil
}

func (s *Server) bind(w *gldap.ResponseWriter, r *gldap.Request) {
	resp := r.NewBindResponse(gldap.WithResponseCode(gldap.ResultInvalidCredentials))
	defer func() { _ = w.Write(resp) }()
	m, err := r.GetSimpleBindMessage()
	if err != nil {
		return
	}
	if m.UserName == s.bindDN && string(m.Password) == s.password {
		resp.SetResultCode(gldap.ResultSuccess)
	}
}

var leafRe = regexp.MustCompile(`\(([A-Za-z]+)=([^()]*)\)`)

// matches evaluates the filter shapes ldapx uses: leaves outside an "(|"
// must all hold, at least one leaf inside it must hold.
func matches(filter string, attrs map[string][]string, dn string) bool {
	all, any := filter, ""
	if i := strings.Index(filter, "(|"); i >= 0 {
		all, any = filter[:i], filter[i:]
	}
	holds := func(attr, val string) bool {
		if strings.EqualFold(attr, "distinguishedName") && strings.EqualFold(val, dn) {
			return true
		}
		if strings.EqualFold(attr, "objectCategory") {
			attr = "objectClass" // the test data has no categories
			if val == "person" {
				val = "user"
			}
		}
		for k, vs := range attrs {
			if !strings.EqualFold(k, attr) {
				continue
			}
			for _, v := range vs {
				if strings.EqualFold(v, val) || wildcard(val, v) {
					return true
				}
			}
		}
		return false
	}
	for _, m := range leafRe.FindAllStringSubmatch(all, -1) {
		if !holds(m[1], m[2]) {
			return false
		}
	}
	if any == "" {
		return true
	}
	for _, m := range leafRe.FindAllStringSubmatch(any, -1) {
		if holds(m[1], m[2]) {
			return true
		}
	}
	return false
}

func (s *Server) search(w *gldap.ResponseWriter, r *gldap.Request) {
	done := r.NewSearchDoneResponse(gldap.WithResponseCode(gldap.ResultSuccess))
	defer func() { _ = w.Write(done) }()
	m, err := r.GetSearchMessage()
	if err != nil {
		done.SetResultCode(gldap.ResultOperationsError)
		return
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	for dn, attrs := range s.entries {
		if !strings.HasSuffix(strings.ToLower(dn), strings.ToLower(m.BaseDN)) || !matches(m.Filter, attrs, dn) {
			continue
		}
		e := r.NewSearchResponseEntry(dn)
		for k, v := range attrs {
			e.AddAttribute(k, v)
		}
		_ = w.Write(e)
	}
}

func (s *Server) modify(w *gldap.ResponseWriter, r *gldap.Request) {
	res := r.NewModifyResponse(gldap.WithResponseCode(gldap.ResultSuccess))
	defer func() { _ = w.Write(res) }()
	m, err := r.GetModifyMessage()
	if err != nil {
		res.SetResultCode(gldap.ResultOperationsError)
		return
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	e, ok := s.entries[m.DN]
	if !ok {
		res.SetResultCode(gldap.ResultNoSuchObject)
		return
	}
	for _, c := range m.Changes {
		name := c.Modification.Type
		vals := make([]string, len(c.Modification.Vals))
		for i, v := range c.Modification.Vals {
			vals[i] = berOctets(v)
		}
		switch c.Operation {
		case gldap.ReplaceAttribute:
			e[name] = append([]string(nil), vals...)
		case gldap.AddAttribute:
			e[name] = append(e[name], vals...)
		case gldap.DeleteAttribute:
			if len(vals) == 0 {
				delete(e, name)
				continue
			}
			var kept []string
			for _, v := range e[name] {
				drop := false
				for _, d := range vals {
					if strings.EqualFold(v, d) {
						drop = true
					}
				}
				if !drop {
					kept = append(kept, v)
				}
			}
			e[name] = kept
		}
	}
}

// testTLS makes a self-signed certificate for 127.0.0.1 and the client side
// that trusts it.
func testTLS(t *testing.T) (server, client *tls.Config) {
	t.Helper()
	key, err := ecdsa.GenerateKey(elliptic.P256(), rand.Reader)
	if err != nil {
		t.Fatal(err)
	}
	tmpl := &x509.Certificate{
		SerialNumber: big.NewInt(1), Subject: pkix.Name{CommonName: "ldaptest"},
		NotBefore: time.Now().Add(-time.Hour), NotAfter: time.Now().Add(time.Hour),
		IPAddresses: []net.IP{net.ParseIP("127.0.0.1")}, IsCA: true, BasicConstraintsValid: true,
		KeyUsage: x509.KeyUsageDigitalSignature | x509.KeyUsageCertSign, ExtKeyUsage: []x509.ExtKeyUsage{x509.ExtKeyUsageServerAuth},
	}
	der, err := x509.CreateCertificate(rand.Reader, tmpl, tmpl, &key.PublicKey, key)
	if err != nil {
		t.Fatal(err)
	}
	cert, _ := x509.ParseCertificate(der)
	pool := x509.NewCertPool()
	pool.AddCert(cert)
	server = &tls.Config{Certificates: []tls.Certificate{{Certificate: [][]byte{der}, PrivateKey: key}}, MinVersion: tls.VersionTLS12}
	client = &tls.Config{RootCAs: pool, ServerName: "127.0.0.1", MinVersion: tls.VersionTLS12}
	return server, client
}

// berOctets strips the BER OCTET STRING header gldap leaves on modify values.
func berOctets(v string) string {
	if len(v) < 2 || v[0] != 0x04 {
		return v
	}
	n := int(v[1])
	switch {
	case n < 0x80 && len(v) == 2+n:
		return v[2:]
	case n == 0x81 && len(v) >= 3 && len(v) == 3+int(v[2]):
		return v[3:]
	case n == 0x82 && len(v) >= 4 && len(v) == 4+(int(v[2])<<8|int(v[3])):
		return v[4:]
	}
	return v
}

// wildcard matches a "*part*" substring filter value.
func wildcard(pattern, v string) bool {
	if !strings.HasPrefix(pattern, "*") || !strings.HasSuffix(pattern, "*") || len(pattern) < 3 {
		return false
	}
	return strings.Contains(strings.ToLower(v), strings.ToLower(pattern[1:len(pattern)-1]))
}
