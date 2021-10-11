package api

import (
	"github.com/tealeg/xlsx"
	"strconv"
)

// https://docs.microsoft.com/en-us/javascript/api/excel/excel.booleancellvalue?view=excel-js-preview

func unwrap(v interface{}, err error) interface{} {
	if err != nil {
		panic(err)
	}
	return v
}

func pIf(b bool) {
	if b {
		panic("")
	}
}

type Value struct {
	text      string
	type_     *ValueType
	primitive interface{}
}

func (v *Value) Type() *ValueType { return v.type_ }
func (v *Value) Boolean() bool {
	pIf(*v.type_ != Boolean)
	return unwrap(strconv.ParseBool(v.text)).(bool)
}
func (v *Value) Float64() float64 {
	pIf(*v.type_ != Double)
	return unwrap(strconv.ParseFloat(v.text, 64)).(float64)
}

// https://docs.microsoft.com/en-us/javascript/api/excel/excel.rangevaluetype?view=excel-js-preview#excel_Excel_RangeValueType_boolean_member
type ValueType = string

const (
	Unknown   ValueType = "Unknown"
	Boolean   ValueType = "Boolean"
	Double    ValueType = "Double"
	Empty     ValueType = "Empty"
	Error     ValueType = "Error"
	Integer   ValueType = "Integer"
	RichValue ValueType = "RichValue"
	String    ValueType = "String"
)

func ValueFromImpl(c *xlsx.Cell) interface{} {
	t := c.Type()
	switch t {
	case xlsx.CellTypeBool:
		return unwrap(strconv.ParseBool(c.Value)).(bool)
	case xlsx.CellTypeString:
		return c.Value
	case xlsx.CellTypeNumeric:
		i, err := strconv.Atoi(c.Value)
		if err != nil {
			return i
		}
		return unwrap(strconv.ParseFloat(c.Value, 64)).(float64)
	case xlsx.CellTypeStringFormula:
		// todo
		return c.Value
	default:
		pIf(true)
	}
	return nil
}

func TypeFromImpl(c *xlsx.Cell) ValueType {
	t := c.Type()
	switch t {
	case xlsx.CellTypeBool:
		return Boolean
	case xlsx.CellTypeString:
		return String
	case xlsx.CellTypeNumeric:
		_, err := strconv.Atoi(c.Value)
		if err != nil {
			return Integer
		}
		return Double
	case xlsx.CellTypeStringFormula:
		// todo
		return String
	default:
		pIf(true)
	}
	return Unknown // todo should panic
}
