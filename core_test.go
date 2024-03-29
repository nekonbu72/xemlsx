package xemlsx

import (
	"path"
	"testing"

	"github.com/nekonbu72/mailg"
	"github.com/nekonbu72/sjson/sjson"
)

type TestExpect struct {
	FileName string `json:"fileName"`
	Sheet    string `json:"sheet"`
	Row      int    `json:"row"`
	Column   int    `json:"column"`
	Value    string `json:"value"`
}

const (
	testDir = "test"
	test    = "test.json"
	expect  = "expect.json"
)

func textExpect() *TestExpect {
	e := new(TestExpect)
	if err := sjson.OpenDecode(path.Join(testDir, expect), e); err != nil {
		panic(err)
	}
	return e
}

func testClient(ci *mailg.ConnInfo) *mailg.Client {
	c, err := mailg.Login(ci)
	if err != nil {
		panic(err)
	}
	return c
}
func TestToXLSX(t *testing.T) {
	st, err := mailg.NewSetting(path.Join(testDir, test))
	if err != nil {
		t.Errorf("NewSetting: %v\n", err)
	}

	c := testClient(st.ConnInfo)
	defer c.Logout()
	e := textExpect()

	done := make(chan interface{})
	defer close(done)
	attachmentStream := c.FetchAttachment(done, st.Criteria)

	var xs []*XLSX
	for x := range ToXLSX(done, attachmentStream) {
		xs = append(xs, x)
	}

	if len(xs) != 1 {
		t.Errorf("len: %v\n", len(xs))
		return
	}

	if xs[0].FileName != e.FileName {
		t.Errorf("FileName: %v\n", xs[0].FileName)
		return
	}

	s, ok := xs[0].Sheet[e.Sheet]
	if !ok {
		t.Errorf("Sheet[%v]: %v\n", e.Sheet, ok)
		return
	}

	v := s.Cell(e.Row, e.Column).String()
	if v != e.Value {
		t.Errorf("Cell(%v, %v): %v\n", e.Row, e.Column, v)
		return
	}

}
