package dump

import (
	"github.com/orpharion/go-excelAPI/pkg/api"
	m "github.com/orpharion/go-excelAPI/pkg/object"
)

func Workbook(workbook *api.Workbook) (w m.Workbook) {
	w.Worksheets = WorksheetCollection(workbook.GetWorksheets())
	return
}

func WorksheetCollection(worksheetCollection *api.WorksheetCollection) (wc m.WorksheetCollection) {
	items := worksheetCollection.Items()
	wc.Items = make([]m.Worksheet, len(items))
	for i, ws := range items {
		wc.Items[i] = Worksheet(ws)
	}
	return
}

func Worksheet(worksheet *api.Worksheet) (ws m.Worksheet) {
	ws.Name = worksheet.GetName()
	ws.Range = Range(worksheet.GetUsedRange())
	return
}

func Range(range_ api.Range) (r m.Range) {
	r.Values = range_.GetValues()
	return
}
