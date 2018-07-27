## Main functions

### pyexcel_export.get_data(in_file, **flags)

- in_file
    - Supported file formats are *.xlsx, *.yaml, *.json and *.pyexcel.json
- flags
    - These are flags for `Meta()`. Supported flags are:-
        - freeze_header (default=True): whether you want to freeze the first row (of non `_meta` sheets)
        - col_width_fit_param_keys (default=True): makes column in `_meta` worksheet wide enough to fit the text.
        - col_width_fit_ids (default=True): makes any column with the header ending with "id" wide enough to fit the id, while retaining the cell type "number"
        - minimum_col_width (default=20): see your own number to enable a minimum bigger column width (or cell width).
        - wrap_text (default=True): enable text wrapping in all cells.
        - bool_as_string (default=True): as a bug fix for openpyxl to rendering boolean in LibreOffice, the boolean becomes `"true"` instead of simply `true`.
        - merge_hidden_tables (default=True): whether then table with the name starting with `_` is rendered altogether in `_meta` worksheet.

### pyexcel_export.get_meta(in_file=None, **flags)

- in_file: if in_file is not specified, or is None, the default `Meta()` is returned.
- flags: same as `pyexcel_export.get_data()`

### pyexcel_export.save_data(out_file, data, meta=None, retain_meta=False, retain_styles=False, **flags)

- out_file: the exporting filename
    - Supported file formats are `*.xlsx`, `*.yaml`, `*.json` and `*.pyexcel.json`.
- data: the data in OrderDict of 2-dimensional array. Same as of `pyexcel_xlsx.get_data()`
- meta: the `Meta()` to format your worksheet. If not specified, the default `Meta()` is used.
- retain_meta (default=False): whether to explicitly show the worksheet `_meta` in your workbook (by default it will be hidden).
- retain_styles (default=False): whether to export the base64 encoded stylesheets to your `_meta` worksheet.
- flags: same as `pyexcel_export.get_data()`

#### Sample file format for \*.yaml and \*.pyexcel.json are also provided in this folder.

## Writing hidden tables (hidden worksheets)

Just name the worksheets starting with `_` and set the flag `allow_table_hiding` to True (default: True). The tables will instead appear at the bottom of `_meta`.

## ExcelLoader class

This is available from `from pyexcel_export import ExcelLoader`

### ExcelLoader(in_file: str=None, **flags)

- in_file (default=None): if in_file isn't specified, it will create ExcelLoader.data with empty data and ExcelLoader.mea with default `Meta()`
- flags: same as `pyexcel_export.get_data()`

### ExcelLoader.save(out_file: str, retain_meta=True, out_format=None, retain_styles=True)

Same as `pyexcel_export.save_data()`

### Properties

```python
>>> from pyexcel_export import ExcelLoader
>>> loader = ExcelLoader("test.xlsx")
>>> loader.formatted_object
OrderedDict([('_meta', [['created', '2018-07-15T05:32:43.976194'], ['modified', '2018-07-16T07:31:26.232350'], ['has_header', True], ['freeze_header', True], ['col_width_fit_param_keys', True], ['col_width_fit_ids', True], ['bool_as_string', True], ['allow_table_hiding', True], ['_styles', <_io.BytesIO object at 0x103d7f678>]]), ('test', [['"id"', '"English"', '"Pinyin"', '"Hanzi"', '"Audio"', '"Tags"'], ['1419644212689', '"Hello!"', '"Nǐ hǎo!"', '"你好！"', '"[sound:tmp1cctcn.mp3]"', '""'], ['1419644212690', '"What are you saying?"', '"Nǐ shuō shénme?"', '"你说什么？"', '"[sound:tmp4tzxbu.mp3]"', '""'], ['1419644212691', '"What did you do?"', '"nǐ zuò le shénme ?"', '"你做了什么？"', '"[sound:333012.mp3]"', '""']])])
>>> loader.data
OrderedDict([('test', [['id', 'English', 'Pinyin', 'Hanzi', 'Audio', 'Tags'], [1419644212689, 'Hello!', 'Nǐ hǎo!', '你好！', '[sound:tmp1cctcn.mp3]', ''], [1419644212690, 'What are you saying?', 'Nǐ shuō shénme?', '你说什么？', '[sound:tmp4tzxbu.mp3]', ''], [1419644212691, 'What did you do?', 'nǐ zuò le shénme ?', '你做了什么？', '[sound:333012.mp3]', '']])])
>>> loader.meta
Meta([
  ('created', '2018-07-15T05:32:43.976194'),
  ('modified', '2018-07-16T07:31:26.232350'),
  ('has_header', True),
  ('freeze_header', True),
  ('col_width_fit_param_keys', True),
  ('col_width_fit_ids', True),
  ('bool_as_string', True),
  ('allow_table_hiding', True),
  ('_styles', <_io.BytesIO object at 0x1065dff10>)
])
>>> loader.in_file
/Users/patarapolw/PycharmProjects/excelformatter/tests/input/test.xlsx
```
