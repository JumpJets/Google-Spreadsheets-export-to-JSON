# Google Spreadsheets export to JSON

This code using [**Google Apps Script**](https://script.google.com/home)

It allow you to convert tabular data into JSON. For example:

| A | B |
|---|---|
| 1 | x |
| 2 | y |

will be:
```json
[
    {
        "a": 1,
        "b": "x"
    },
    {
        "a": 2,
        "b": "y"
    }
]
```
Name of the columns would be lowercased and all non-alphanumerical letters will be replaced to `_`. Like `Some Column` would become `some_column`.
The choice of the using snake case notation is primary for the SQL parity, where table name like `someColumn`, `SomeColumn` would be considered as `somecolumn` unless enquoted with `"someColumn"`.

Be sure that empty cells are omitted.

At this moment there is no configuration, but it is possible to add if it would be needed.

## Usage

1. Open your Sheet
2. Click in menu `Extensions` → `Apps Script`
3. New tab will be open with empty script
4. Copy content of the file `export_to_json.gs` into this Apps Script text editor
5. Rename project name (on top) and script file (`.gs`) to a meaningful name.
6. Click Save button 💾
7. Click Execute button ▶
8. For the first time you would be asked for permissions to run unknown code, you would need to accept and check checkboxes to allow permission.

If everything is ok you will see logs that script is successfully executed. You can close this window and go back to your Sheet.
You should be able to see in menu new option `Export JSON`.

<img width="827" height="102" alt="menu button" src="https://github.com/user-attachments/assets/5843eddb-17ad-4be7-acec-00d69016cac6" />

Because there is limited permissions, saving it as a file would not be possible, but you will see a popup dialog where you would be able to easily select and copy with `Ctrl`+`C`. And then paste into an empty file locally.

## Bugs

If you find any bugs or suggestions, feel free to open an issue, but make sure to provide sample data.
