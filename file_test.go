package xemlsx_test

import (
	"testing"
	"time"

	"github.com/nekonbu72/mailg"
	"github.com/nekonbu72/sjson/sjson"
	"github.com/nekonbu72/xemlsx"
)

type MyTest struct {
	Host     string `json:"host"`
	Port     string `json:"port"`
	User     string `json:"user"`
	Password string `json:"password"`

	TimeFormat   string `json:"timeFormat"`
	SinceDay     string `json:"sinceDay"`
	DaysDuration int    `json:"daysDuration"`

	Name string `json:"name"`

	Sheet  string `json:"sheet"`
	Row    int    `json:"row"`
	Column int    `json:"column"`
	Value  string `json:"value"`

	Criteria *mailg.Criteria
}

const (
	jsonpath string = "test.json"
)

func createMyTest() *MyTest {
	mt := new(MyTest)
	if err := sjson.OpenDecode(jsonpath, mt); err != nil {
		panic("")
	}
	since, _ := time.Parse(mt.TimeFormat, mt.SinceDay)
	before := since.AddDate(0, 0, mt.DaysDuration)
	mt.Criteria = &mailg.Criteria{Since: since, Before: before}
	return mt
}

func TestOpenAttachment(t *testing.T) {

	mt := createMyTest()

	c, err := mailg.Login(
		&mailg.ConnInfo{
			Host:     mt.Host,
			Port:     mt.Port,
			User:     mt.User,
			Password: mt.Password,
		},
	)
	defer c.Logout()
	if err != nil {
		t.Errorf("CreateClient: %v\n", err)
		return
	}

	done := make(chan interface{})
	defer close(done)
	attachmentStream := c.FetchAttachment(done, mt.Name, mt.Criteria)

	for a := range attachmentStream {
		x, err := xemlsx.OpenAttachment(a)
		if err != nil {
			t.Errorf("OpenAttachment: %v\n", err)
			return
		}

		s, ok := x.Sheet[mt.Sheet]
		if !ok {
			t.Errorf("Sheet: %v\n", ok)
			return
		}

		v := s.Cell(mt.Row, mt.Column).String()
		if v != mt.Value {
			t.Errorf("Cell: %v\n", v)
			return
		}
	}
}
