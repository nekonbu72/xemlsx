package xemlsx

import (
	"bytes"
	"io"
	"log"

	"github.com/nekonbu72/mailg"
	"github.com/tealeg/xlsx"
)

const (
	errLimit = 3
)

type XLSX struct {
	FileName string
	*xlsx.File
}

func toXLSX(a *mailg.Attachment) (*XLSX, error) {
	r, size, err := readerAt(a)
	if err != nil {
		return nil, err
	}

	f, err := xlsx.OpenReaderAt(r, size)
	if err != nil {
		return nil, err
	}

	return &XLSX{FileName: a.FileName, File: f}, nil
}

func readerAt(r io.Reader) (io.ReaderAt, int64, error) {
	buf := bytes.NewBuffer(nil)
	size, err := buf.ReadFrom(r)
	if err != nil {
		return nil, 0, err
	}
	return bytes.NewReader(buf.Bytes()), size, nil
}

func ToXLSX(
	done <-chan interface{},
	attachmentStream <-chan *mailg.Attachment,
) <-chan *XLSX {
	xlsxStream := make(chan *XLSX)
	go func() {
		defer close(xlsxStream)

		errCount := 0
		for a := range attachmentStream {
			x, err := toXLSX(a)
			if err != nil {
				log.Printf("toXLSX: %v\n", err)
				errCount++
				if errCount >= errLimit {
					log.Println("To many errors, breaking!")
					break
				}
				continue
			}

			select {
			case <-done:
				return
			case xlsxStream <- x:
			}
		}
	}()
	return xlsxStream
}
