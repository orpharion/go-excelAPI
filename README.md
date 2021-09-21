# excelAPI

## Overview

`excelAPI` provides an API and Object model similar to Excel's
new [JavaScript API](https://docs.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview), built on the
excellent [xlsx](https://github.com/tealeg/xlsx) library.
It doesn't require Excel.

It's primary purpose is transcoding excel workbooks to neutral data formats, and doesn't support modifying workbooks.

It offers two libraries:

- [`pkg/api`](pkg/api/doc.go): Javascript-like API, providing Object-Oriented model.
- [`pkg/object`](pkg/object/doc.go): Javascript-like data model, derived from `./pkg/api`, intended to dump data as follows:
  
  ``` js.JSON.stringify(MSworkbookJSObject) ~= go.json.Marshal(excelAPIworkbookStruct)```

One command:

- [`cmd/go-excelAPI`] - dump excel workbooks to JSON.

The following API's:

- [`api/cue`](api/cue) - CUE bindings for object model (not the api)

The data structures differ between these packages, reflecting their different intents.

## Usage

```shell
go-excelAPI toJson workbook.xlsx > workbook.json
```

## Direction

1. Derive `api/interfaces` and `object/model` directly from the reference documentation.
2. Support writing by providing appropriate structures.

