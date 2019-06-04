package xemlsx

import (
	"bytes"
	"io"

	"github.com/nekonbu72/mailg"
	"github.com/tealeg/xlsx"
)

type XLSX struct {
	*mailg.Attachment
	*xlsx.File
}

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

func AttachmentToXLSX(
	done <-chan interface{},
	attachmentStream <-chan *mailg.Attachment,
) <-chan *XLSX {
	xlsxStream := make(chan *XLSX)
	go func() {
		defer close(xlsxStream)
		var (
			x   *XLSX
			err error
		)
		for a := range attachmentStream {
			select {
			case <-done:
				return
			case xlsxStream <- x:

			default:
				x, err = openAttachment(a)
			}
		}
	}()
}
