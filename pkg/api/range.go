// Package range_ implements basic Range functionality.
// See https://docs.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview
package api

type Address = string

// Range represents a set of one or more contiguous cells such as a cell, a row, a column, or a block of cells.
// https://docs.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview\
// Currently, only rectangular, single-set ranges are supported.
type Range struct {
	address     Address
	columnCount uint
	columnIndex uint
	rowCount    uint
	rowIndex    uint
	// Represents the raw values of the specified range. The data returned could be a string,
	// number, or boolean. Cells that contain an error will return the error string. If the
	// returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets
	// this value as a formula.
	// https://docs.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview#values
	values     [][]interface{}
	valueTypes [][]ValueType
	worksheet  *Worksheet
}

func ByWorksheetAndIndexes(worksheet *Worksheet, startRow, startColumn, rowCount, columnCount uint) Range {
	rows := worksheet.Impl().Rows[startRow : startRow+rowCount]
	values := make([][]interface{}, rowCount)
	valueTypes := make([][]ValueType, rowCount)
	for iR, row := range rows {
		values[iR] = make([]interface{}, columnCount)
		valueTypes[iR] = make([]ValueType, columnCount)
		_r := row.Cells[startColumn:columnCount]
		for iC, c := range _r {
			values[iR][iC] = ValueFromImpl(c)
			valueTypes[iR][iC] = TypeFromImpl(c)
		}
	}
	return Range{
		"",
		columnCount,
		startColumn,
		rowCount,
		startRow,
		values,
		valueTypes,
		worksheet,
	}
}

func (r *Range) GetAddress() Address             { return r.address }
func (r *Range) GetColumnCount() uint            { return r.columnCount }
func (r *Range) GetColumnIndex() uint            { return r.columnIndex }
func (r *Range) GetRowIndex() uint               { return r.rowIndex }
func (r *Range) GetValues() [][]interface{}      { return r.values }
func (r *Range) GetValueTypes() [][]ValueType { return r.valueTypes }

//func (r *Range) SetValues(v [][]interface{}) { r.values = v }
