package worker

import (
	"crypto/hmac"
	"crypto/rand"
	"crypto/sha256"
	"encoding/hex"
	"strings"
	"sync"
	"time"
)

// Pairing: the worker shows a one-time code; the client proves it knows the
// code over the certificate fingerprints it saw in the TLS handshake, and
// the worker answers with a proof the same way. A machine in the middle
// shows a different certificate to one side, so the proofs do not match.

const (
	codeAlphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789" // no 0/O, 1/I
	codeLength   = 16 // 80 bits
	codeTTL      = 10 * time.Minute
	maxAttempts  = 5  // per address
	maxTotal     = 20 // all addresses together
)

// NewCode returns a pairing code like "K7QX-M2PA-9DHT-R4WE".
func NewCode() string {
	b := make([]byte, codeLength)
	_, _ = rand.Read(b)
	out := make([]byte, 0, codeLength+2)
	for i, v := range b {
		if i > 0 && i%4 == 0 {
			out = append(out, '-')
		}
		out = append(out, codeAlphabet[int(v)%len(codeAlphabet)])
	}
	return string(out)
}

func normalizeCode(c string) string {
	return strings.ToUpper(strings.NewReplacer("-", "", " ", "").Replace(c))
}

// proof binds the code to the two fingerprints, in order (who proves first).
func proof(code, label, fpA, fpB string) string {
	m := hmac.New(sha256.New, []byte(normalizeCode(code)))
	m.Write([]byte(label + "\x00" + fpA + "\x00" + fpB))
	return hex.EncodeToString(m.Sum(nil))
}

// ClientProof is sent by the client: it has seen workerFP, presents clientFP.
func ClientProof(code, clientFP, workerFP string) string {
	return proof(code, "client", clientFP, workerFP)
}

// WorkerProof is returned by the worker.
func WorkerProof(code, clientFP, workerFP string) string {
	return proof(code, "worker", clientFP, workerFP)
}

// pairing is a code waiting to be used once.
type pairing struct {
	mu       sync.Mutex
	code     string
	expires  time.Time
	attempts int            // all addresses
	perIP    map[string]int // a noisy neighbour cannot use up the code for others
	used     bool
}

func newPairing(code string) *pairing {
	return &pairing{code: code, expires: time.Now().Add(codeTTL), perIP: map[string]int{}}
}

// active reports whether the code can still be used.
func (p *pairing) active() bool {
	if p == nil {
		return false
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	return !p.used && p.attempts < maxTotal && time.Now().Before(p.expires)
}

// check consumes an attempt; a match uses the code up.
func (p *pairing) check(ip, clientFP, workerFP, got string) (string, bool) {
	p.mu.Lock()
	defer p.mu.Unlock()
	if p.used || p.attempts >= maxTotal || p.perIP[ip] >= maxAttempts || time.Now().After(p.expires) {
		return "", false
	}
	p.attempts++
	p.perIP[ip]++
	if !hmac.Equal([]byte(got), []byte(ClientProof(p.code, clientFP, workerFP))) {
		return "", false
	}
	p.used = true
	return WorkerProof(p.code, clientFP, workerFP), true
}
