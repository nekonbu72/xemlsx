package xemlsx

import (
	"bytes"
	"io"
	"log"

	"github.com/nekonbu72/mailg"
	"github.com/tealeg/xlsx"
)

type XLSX struct {
	*mailg.Attachment
	*xlsx.File
}

const (
	errLimit int = 3
)

func openAttachment(a *mailg.Attachment) (*XLSX, error) {
	r, size, err := readerAt(a)
	if err != nil {
		return nil, err
	}

	f, err := xlsx.OpenReaderAt(r, size)
	if err != nil {
		return nil, err
	}

	return &XLSX{Attachment: a, File: f}, nil
}

func readerAt(r io.Reader) (io.ReaderAt, int64, error) {
	buf := bytes.NewBuffer(nil)
	size, err := buf.ReadFrom(r)
	if err != nil {
		return nil, 0, err
	}
	return bytes.NewReader(buf.Bytes()), size, nil
}

type result struct {
	Error error
	XLSX  *XLSX
}

func toXLSX(
	done <-chan interface{},
	attachmentStream <-chan *mailg.Attachment,
) <-chan result {
	results := make(chan result)
	go func() {
		defer close(results)

		for a := range attachmentStream {
			x, err := openAttachment(a)
			r := result{Error: err, XLSX: x}
			select {
			case <-done:
				return
			case results <- r:
			}
		}
	}()
	return results
}

func resultFilter(
	done <-chan interface{},
	results <-chan result,
	errLimit int,
) <-chan *XLSX {
	xlsxStream := make(chan *XLSX)
	go func() {
		defer close(xlsxStream)

		errCount := 0
		for r := range results {
			if r.Error != nil {
				log.Printf("error: %v", r.Error)
				errCount++
				if errCount >= errLimit {
					log.Println("Too many errors, breaking!")
					break
				}
				continue
			}
			select {
			case <-done:
				return
			case xlsxStream <- r.XLSX:
			}
		}
	}()
	return xlsxStream
}

func ToXLSX(
	done <-chan interface{},
	attachmentStream <-chan *mailg.Attachment,
) <-chan *XLSX {
	ch := toXLSX(done, attachmentStream)
	return resultFilter(done, ch, errLimit)
}
