package xlsxParser

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

type XlsxParser struct {
	workbox    *xlsx.File
	outputPath string
	fileName   string
	fileCount  int
	tableHead  *xlsx.Row
	tableBody  []*xlsx.Row
}

func New(fileName, outputPath string) *XlsxParser {
	workbox, err := xlsx.OpenFile(fileName)
	if err != nil {
		panic(err)
	}

	return &XlsxParser{
		workbox:    workbox,
		outputPath: outputPath,
		fileName:   fileName,
	}
}

func (xlsxParser *XlsxParser) ReadTable(rows int, offset int) {
	for _, sheet := range xlsxParser.workbox.Sheets[offset:] {
		sheet.ForEachRow(func(r *xlsx.Row) error {
			if r.GetCoordinate() == 0 {
				xlsxParser.tableHead = r
			} else {
				xlsxParser.tableBody = append(xlsxParser.tableBody, r)
			}

			if len(xlsxParser.tableBody) == rows {
				xlsxParser.createXlsxFile(r.Sheet.Name)
			}

			return nil
		})
	}

	if len(xlsxParser.tableBody) > 0 {
		xlsxParser.createXlsxFile("Page 1")
	}
}

func (xlsxParser *XlsxParser) createXlsxFile(sheetName string) {
	xlsxParser.fileCount++

	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet(sheetName)
	if err != nil {
		panic(err)
	}

	head := sheet.AddRow()
	xlsxParser.tableHead.ForEachCell(func(c *xlsx.Cell) error {
		head.PushCell(c)
		return nil
	})

	for _, row := range xlsxParser.tableBody {
		body := sheet.AddRow()
		row.ForEachCell(func(c *xlsx.Cell) error {
			body.PushCell(c)
			return nil
		})
	}

	fileName := strings.Replace(xlsxParser.fileName, ".xlsx", "", 1) + "-" + fmt.Sprint(xlsxParser.fileCount) + ".xlsx"
	filePath := filepath.Join(xlsxParser.outputPath, fileName)
	fmt.Println("Creating file:", filePath)
	err = xlsxFile.Save(filePath)
	if err != nil {
		panic(err)
	}

	xlsxParser.tableBody = nil
}
