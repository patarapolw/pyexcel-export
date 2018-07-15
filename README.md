# pyexcel-export - Keep optimal formatting when exported to xlsx format

pyexcel-export is a wrapper around [pyexcel](https://github.com/pyexcel/pyexcel), [pyexcel-xlsx](https://github.com/pyexcel/pyexcel-xlsx) and [openpyxl](https://bitbucket.org/openpyxl/openpyxl) to read the formatting (stylesheets) and update the pre-existing file without destroying the stylesheets.

pyexcel-export also introduces a new exporting format, `*.pyexcel.json` which is based on [NoIndentEncoder](https://stackoverflow.com/a/25935321/9023855). This allows you to edit the spreadsheet in you favorite text editor, without being frustrated by automatically collapsed cells in Excel.

## Known constraints

The "stylesheets" exported from Excel is in a very long base64 encoded format when exported to `*.json` or `*.pyexcel.json`, so exporting is disabled by default.

As stylesheets copying works by [`openpyxl.worksheet.copier.WorksheetCopy`](https://openpyxl.readthedocs.io/en/2.5/_modules/openpyxl/worksheet/copier.html), it may work only on values, styles, dimensions and merged cells, but charts might not be supported.

## Installation

You can install pyexcel-export via pip:

```commandline
$ pip install pyexcel-export
```

or clone it and install it.

## Usage

### Exporting stylesheets

```pydocstring
>>> from pyexcel_export import get_stylesheet
>>> get_stylesheet()
{
  "created": "'2018-07-15T08:03:17.762214'",
  "has_header": "True",
  "freeze_header": "True",
  "col_width_fit_param_keys": "True",
  "col_width_fit_ids": "True"
}
>>> get_stylesheet('test.xlsx')
{
  "created": "'2018-07-12T15:21:25.777499'",
  "modified": "'2018-07-12T15:21:25.777523'",
  "has_header": "True",
  "freeze_header": "True",
  "col_width_fit_param_keys": "True",
  "col_width_fit_ids": "True",
  "_styles": "<_io.BytesIO object at 0x10553b678>"
}
```
### Saving files, while preserving the formatting.
```pydocstring
>>> from pyexcel_export import get_stylesheet, save_data
>>> stylesheet = get_stylesheet('test.xlsx')
>>> data = OrderedDict()
>>> data.update({"Sheet 1": [["header A", "header B", "header C"], [1, 2, 3]]})
>>> save_data("your_file.xlsx", data)
```
### \*.pyexcel.json format
```json
{
  "_meta": [
    ["\"created\"", "\"2018-07-12T15:21:25.777499\""],
    ["\"modified\"", "\"2018-07-15T06:59:22.707162\""],
    ["\"has_header\"", "true"],
    ["\"freeze_header\"", "true"],
    ["\"col_width_fit_param_keys\"", "true"],
    ["\"col_width_fit_ids\"", "true"]
  ],
  "test": [
    ["\"id\"", "\"English\"", "\"Pinyin\"", "\"Hanzi\"", "\"Audio\"", "\"Tags\""],
    ["1419644212689", "\"Hello!\"", "\"Nǐ hǎo!\"", "\"你好！\"", "\"[sound:tmp1cctcn.mp3]\"", "\"\""],
    ["1419644212690", "\"What are you saying?\"", "\"Nǐ shuō shénme?\"", "\"你说什么？\"", "\"[sound:tmp4tzxbu.mp3]\"", "\"\""],
    ["1419644212691", "\"What did you do?\"", "\"nǐ zuò le shénme ?\"", "\"你做了什么？\"", "\"[sound:333012.mp3]\"", "\"\""]
  ]
}
```
