package api

import (
	"github.com/tealeg/xlsx"
)

type Workbook struct {
	file       *xlsx.File // underlying xlsx file
	worksheets WorksheetCollection
}

func (w *Workbook) GetWorksheets() *WorksheetCollection { return &w.worksheets }

// CreateWorkbook - Note: unlike reference, source data isn't optional.
// https://docs.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview#Excel_createWorkbook_base64_
func CreateWorkbook(src []byte) (wb Workbook, err error) {
	file, err := xlsx.OpenBinary(src)
	if err != nil {
		return
	}
	wb.file = file
	wb.worksheets = FromWorkbookImpl(file)
	return
}
