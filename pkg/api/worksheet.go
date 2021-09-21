package api

import (
	"github.com/tealeg/xlsx"
)

// Worksheet - see
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview
type Worksheet struct {
	name  string
	sheet *xlsx.Sheet
}

func (w *Worksheet) GetName() string { return w.name }

// GetRangeByIndexes - see
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.13#getRangeByIndexes_startRow__startColumn__rowCount__columnCount_
func (w *Worksheet) GetRangeByIndexes(startRow, startColumn, rowCount, columnCount uint) Range {
	return ByWorksheetAndIndexes(w, startRow, startColumn, rowCount, columnCount)
}

// GetUsedRange - see
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview#getRange_address_
// todo - no valuesOnly option.
// todo - starting at 0, 0 by default.
func (w *Worksheet) GetUsedRange() Range {
	return ByWorksheetAndIndexes(w, 0, 0, uint(w.sheet.MaxRow), uint(w.sheet.MaxCol))
}

func (w *Worksheet) Impl() *xlsx.Sheet { return w.sheet }

func FromSheetImpl(sheet *xlsx.Sheet) Worksheet {
	return Worksheet{name: sheet.Name, sheet: sheet}
}

// WorksheetCollection of Worksheet. See -
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview
type WorksheetCollection struct {
	items []*Worksheet
}

// FromWorkbookImpl creates a new WorksheetCollection from *xlsx.File.
func FromWorkbookImpl(wb *xlsx.File) (wsc WorksheetCollection) {
	ws_ := wb.Sheets
	wsci := make([]*Worksheet, len(ws_))
	for i, sht_ := range ws_ {
		sht := FromSheetImpl(sht_)
		wsci[i] = &sht
	}
	wsc.items = wsci
	return
}

func (c *WorksheetCollection) Items() []*Worksheet { return c.items }
