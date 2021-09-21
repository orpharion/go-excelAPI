package interfaces

import "github.com/tealeg/xlsx"

type Address = string

type Value interface {
	Boolean() bool
	Float64() float64
}
type ValueType = string

type Range interface {
	GetAddress() Address
	GetColumnCount() uint
	GetColumnIndex() uint
	GetRowIndex() uint
	GetValues() [][]interface{}
	GetValueTypes() [][]ValueType
}

type Worksheet interface {
	GetName() string
	GetRangeByIndexes(startRow, startColumn, rowCount, columnCount uint) Range
	Impl() *xlsx.Sheet
}
