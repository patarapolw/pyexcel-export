## Main functions

### pyexcel_export.get_data(in_file, **flags)

- in_file
    - Supported file formats are *.xlsx, *.yaml, *.json and *.pyexcel.json
- flags
    - These are flags for `Meta()`. Supported flags are:-
        - freeze_header (default=True): whether you want to freeze the first row (of non `_meta` sheets)
        - col_width_fit_param_keys (default=True): makes column in `_meta` worksheet wide enough to fit the text.
        - col_width_fit_ids (default=True): makes any column with the header ending with "id" wide enough to fit the id, while retaining the cell type "number"
        - bool_as_string (default=True): as a bug fix for openpyxl to rendering boolean in LibreOffice, the boolean becomes `"true"` instead of simply `true`.
        - allow_table_hiding (default=True): whether then table with the name starting with `_` is rendered in `_meta` worksheet.

### pyexcel_export.get_meta(in_file=None, **flags)

- in_file: if in_file is not specified, or is None, the default `Meta()` is returned.
- flags: same as `pyexcel_export.get_data()`

### pyexcel_export.save_data(out_file, data, meta=None, retain_meta=False, retain_styles=False, **flags)

- out_file: the exporting filename
    - Supported file formats are *.xlsx, *.yaml, *.json and *.pyexcel.json
- data: the data in OrderDict of 2-dimensional array. Same as of `pyexcel_xlsx.get_data()`
- meta: the `Meta()` to format your worksheet. If not specified, the default `Meta()` is used.
- retain_meta (default=False): whether to retain the worksheet `_meta` in your workbook for future syncing.
- retain_styles (default=False): whether to export the base64 encoded stylesheets to your `_meta` worksheet.
- flags: same as `pyexcel_export.get_data()`

#### Sample file format for \*.yaml and \*.pyexcel.json are also provided in this folder.

## Writing hidden tables (hidden worksheets)

Just name the worksheets starting with `_` and set the flag `allow_table_hiding` to True (default: True).
