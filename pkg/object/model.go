package object

// Workbook see workbook.Workbook.
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-preview#worksheets
type Workbook struct {
	Worksheets WorksheetCollection `json:"worksheets"`
}

// WorksheetCollection see worksheet.Collection.
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview
type WorksheetCollection struct {
	Items Worksheets `json:"items"`
}

type Worksheets []Worksheet

type Address string


// Worksheet see worksheet.Worksheet.
type Worksheet struct {
	Name string `json:"name,omitempty"`
	// Alias for getRange. The entire UsedRange.
	Range Range `json:"ranges,omitempty"`
}

type Range struct {
	Values [][]interface{} `json:"values"`
}
