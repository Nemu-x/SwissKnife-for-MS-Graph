// Package ldapx talks to on-premises Active Directory over LDAP: a simple
// bind with a service account, one connection per operation (no idle
// connection to time out), and only the few operations the catalog needs.
package ldapx

import (
	"context"
	"crypto/tls"
	"crypto/x509"
	"encoding/binary"
	"errors"
	"fmt"
	"net"
	"os"
	"strconv"
	"strings"
	"time"
	"unicode/utf16"

	"github.com/go-ldap/ldap/v3"
)

// TLSMode is how the connection is protected.
type TLSMode string

const (
	LDAPS    TLSMode = "ldaps"    // TLS from the first byte (port 636)
	StartTLS TLSMode = "starttls" // plain connect, then upgrade (port 389)
	Plain    TLSMode = "none"     // labs only: no password resets
)

// Config is an on-prem connection without its password (that lives in the
// OS keychain).
type Config struct {
	ID     string  `json:"id"`
	Name   string  `json:"name"`
	Host   string  `json:"host"`
	Port   int     `json:"port"`
	TLS    TLSMode `json:"tls"`
	BaseDN string  `json:"baseDn"`
	BindDN string  `json:"bindDn"` // DN or user@domain
	// CAFile is an optional PEM file with the domain's CA, for directories
	// whose certificate the OS does not trust.
	CAFile string `json:"caFile,omitempty"`
}

// Validate checks a config before it is saved or used.
func (c Config) Validate() error {
	switch {
	case strings.TrimSpace(c.Host) == "":
		return errors.New("host is required")
	case strings.TrimSpace(c.BaseDN) == "":
		return errors.New("base DN is required")
	case strings.TrimSpace(c.BindDN) == "":
		return errors.New("bind account is required")
	}
	switch c.TLS {
	case LDAPS, StartTLS, Plain:
	default:
		return errors.New("TLS mode must be ldaps, starttls or none")
	}
	if c.Port < 0 || c.Port > 65535 {
		return errors.New("invalid port")
	}
	return nil
}

func (c Config) port() int {
	if c.Port > 0 {
		return c.Port
	}
	if c.TLS == LDAPS {
		return 636
	}
	return 389
}

// Client runs operations against one directory.
type Client struct {
	cfg      Config
	password string
	// TLSConfig overrides the TLS settings (tests with their own CA).
	TLSConfig *tls.Config
}

// New builds a client; nothing is dialled until the first operation.
func New(cfg Config, password string) *Client { return &Client{cfg: cfg, password: password} }

// Config returns the connection settings.
func (c *Client) Config() Config { return c.cfg }

// Secure reports whether the connection is encrypted (password resets
// require it: AD refuses unicodePwd over plain LDAP).
func (c *Client) Secure() bool { return c.cfg.TLS == LDAPS || c.cfg.TLS == StartTLS }

func (c *Client) tlsConfig() (*tls.Config, error) {
	if c.TLSConfig != nil {
		return c.TLSConfig, nil
	}
	cfg := &tls.Config{ServerName: c.cfg.Host, MinVersion: tls.VersionTLS12}
	if c.cfg.CAFile != "" {
		pem, err := os.ReadFile(c.cfg.CAFile)
		if err != nil {
			return nil, fmt.Errorf("CA file: %w", err)
		}
		pool, err := x509.SystemCertPool()
		if err != nil || pool == nil {
			pool = x509.NewCertPool()
		}
		if !pool.AppendCertsFromPEM(pem) {
			return nil, errors.New("CA file: no PEM certificate found")
		}
		cfg.RootCAs = pool
	}
	return cfg, nil
}

// dial connects and binds.
func (c *Client) dial(ctx context.Context) (*ldap.Conn, error) {
	addr := net.JoinHostPort(c.cfg.Host, strconv.Itoa(c.cfg.port()))
	tlsCfg, err := c.tlsConfig()
	if err != nil {
		return nil, err
	}
	dialer := &net.Dialer{Timeout: 10 * time.Second}
	var conn *ldap.Conn
	switch c.cfg.TLS {
	case LDAPS:
		conn, err = ldap.DialURL("ldaps://"+addr, ldap.DialWithDialer(dialer), ldap.DialWithTLSConfig(tlsCfg))
	default:
		conn, err = ldap.DialURL("ldap://"+addr, ldap.DialWithDialer(dialer))
		if err == nil && c.cfg.TLS == StartTLS {
			if err = conn.StartTLS(tlsCfg); err != nil {
				_ = conn.Close()
				return nil, fmt.Errorf("StartTLS: %w", err)
			}
		}
	}
	if err != nil {
		return nil, err
	}
	conn.SetTimeout(30 * time.Second)
	if dl, ok := ctx.Deadline(); ok {
		conn.SetTimeout(time.Until(dl))
	}
	if err := conn.Bind(c.cfg.BindDN, c.password); err != nil {
		_ = conn.Close()
		return nil, fmt.Errorf("bind: %w", err)
	}
	return conn, nil
}

// Test connects and binds.
func (c *Client) Test(ctx context.Context) error {
	conn, err := c.dial(ctx)
	if err != nil {
		return err
	}
	_ = conn.Close()
	return nil
}

// Entry is a directory object: its DN and attributes.
type Entry struct {
	DN    string              `json:"dn"`
	Attrs map[string][]string `json:"attrs"`
}

// Get returns an attribute's first value ("" if absent).
func (e Entry) Get(name string) string {
	for k, v := range e.Attrs {
		if strings.EqualFold(k, name) && len(v) > 0 {
			return v[0]
		}
	}
	return ""
}

// All returns every value of an attribute.
func (e Entry) All(name string) []string {
	for k, v := range e.Attrs {
		if strings.EqualFold(k, name) {
			return v
		}
	}
	return nil
}

// Search runs a subtree search under the base DN.
func (c *Client) Search(ctx context.Context, filter string, attrs []string, limit int) ([]Entry, error) {
	conn, err := c.dial(ctx)
	if err != nil {
		return nil, err
	}
	defer func() { _ = conn.Close() }()
	req := ldap.NewSearchRequest(c.cfg.BaseDN, ldap.ScopeWholeSubtree, ldap.NeverDerefAliases, limit, 30, false, filter, attrs, nil)
	res, err := conn.Search(req)
	if err != nil && !ldap.IsErrorWithCode(err, ldap.LDAPResultSizeLimitExceeded) {
		return nil, err
	}
	if res == nil {
		return nil, err
	}
	out := make([]Entry, 0, len(res.Entries))
	for _, e := range res.Entries {
		m := map[string][]string{}
		for _, a := range e.Attributes {
			m[a.Name] = a.Values
		}
		out = append(out, Entry{DN: e.DN, Attrs: m})
	}
	return out, nil
}

// UserAttrs are read for every user.
var UserAttrs = []string{"distinguishedName", "sAMAccountName", "userPrincipalName", "displayName", "mail",
	"userAccountControl", "lockoutTime", "pwdLastSet", "lastLogonTimestamp", "memberOf"}

// FindUser resolves a sAMAccountName, UPN, mail or DN to exactly one user.
func (c *Client) FindUser(ctx context.Context, q string) (Entry, error) {
	q = strings.TrimSpace(q)
	if q == "" {
		return Entry{}, errors.New("user is required")
	}
	v := ldap.EscapeFilter(q)
	filter := fmt.Sprintf("(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=%s)(userPrincipalName=%s)(mail=%s)(distinguishedName=%s)))", v, v, v, v)
	return c.one(ctx, filter, UserAttrs, "user", q)
}

// FindGroup resolves a group by name, sAMAccountName or DN.
func (c *Client) FindGroup(ctx context.Context, q string) (Entry, error) {
	q = strings.TrimSpace(q)
	if q == "" {
		return Entry{}, errors.New("group is required")
	}
	v := ldap.EscapeFilter(q)
	filter := fmt.Sprintf("(&(objectClass=group)(|(cn=%s)(sAMAccountName=%s)(distinguishedName=%s)))", v, v, v)
	return c.one(ctx, filter, []string{"distinguishedName", "cn", "sAMAccountName", "member", "groupType"}, "group", q)
}

func (c *Client) one(ctx context.Context, filter string, attrs []string, kind, q string) (Entry, error) {
	list, err := c.Search(ctx, filter, attrs, 2)
	if err != nil {
		return Entry{}, err
	}
	switch len(list) {
	case 0:
		return Entry{}, fmt.Errorf("%s %q not found", kind, q)
	case 1:
		return list[0], nil
	default:
		return Entry{}, fmt.Errorf("%q matches more than one %s — use the distinguished name", q, kind)
	}
}

// Mod is one attribute change.
type Mod struct {
	Op     string // replace | add | delete
	Attr   string
	Values []string
}

// Modify applies changes to one object.
func (c *Client) Modify(ctx context.Context, dn string, mods ...Mod) error {
	conn, err := c.dial(ctx)
	if err != nil {
		return err
	}
	defer func() { _ = conn.Close() }()
	req := ldap.NewModifyRequest(dn, nil)
	for _, m := range mods {
		switch m.Op {
		case "replace":
			req.Replace(m.Attr, m.Values)
		case "add":
			req.Add(m.Attr, m.Values)
		case "delete":
			req.Delete(m.Attr, m.Values)
		default:
			return fmt.Errorf("unknown modify op %q", m.Op)
		}
	}
	return conn.Modify(req)
}

// ErrNeedsTLS refuses a password reset over plain LDAP.
var ErrNeedsTLS = errors.New("password resets require LDAPS or StartTLS")

// SetPassword sets a user's password (AD unicodePwd: the quoted password in
// UTF-16LE) and, if asked, makes the user change it at next sign-in.
func (c *Client) SetPassword(ctx context.Context, dn, password string, mustChange bool) error {
	if !c.Secure() {
		return ErrNeedsTLS
	}
	mods := []Mod{{Op: "replace", Attr: "unicodePwd", Values: []string{UnicodePwd(password)}}}
	if mustChange {
		mods = append(mods, Mod{Op: "replace", Attr: "pwdLastSet", Values: []string{"0"}})
	}
	return c.Modify(ctx, dn, mods...)
}

// UnicodePwd encodes a password the way AD expects it in unicodePwd.
func UnicodePwd(pw string) string {
	u := utf16.Encode([]rune(`"` + pw + `"`))
	b := make([]byte, 2*len(u))
	for i, r := range u {
		binary.LittleEndian.PutUint16(b[2*i:], r)
	}
	return string(b)
}

// userAccountControl flags.
const (
	UACDisabled = 0x2
)

// UAC returns the user's userAccountControl value.
func UAC(e Entry) int {
	n, _ := strconv.Atoi(e.Get("userAccountControl"))
	return n
}

// Locked reports a recorded lockout (lockoutTime > 0). AD clears it lazily
// after the lockout duration, so this can show a lockout that has expired.
func Locked(e Entry) bool {
	n, _ := strconv.ParseInt(e.Get("lockoutTime"), 10, 64)
	return n > 0
}

// FileTime converts an AD timestamp (100ns since 1601) to time ("zero" if unset).
func FileTime(v string) time.Time {
	n, err := strconv.ParseInt(v, 10, 64)
	if err != nil || n <= 0 || n == 0x7FFFFFFFFFFFFFFF {
		return time.Time{}
	}
	return time.Unix(0, (n-116444736000000000)*100).UTC()
}
