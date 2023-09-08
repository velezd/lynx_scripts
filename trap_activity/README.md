# Trap activity

Script for counting trap active days per season per year.

## Dependencies

 - python3-openpyxl
 - python3-tkinter

## Usage

The script works with xslx files in format:

```
Lokalita | ID lokality |  Oblast  | GPS - šířka | GPS - délka | Model fotopasti |  Správce | {day}.{month}.{year} +
anything |   anything  | anything |   anything  |   anything  |    anything     | anything | 1 or 0 +
```

First line has to be the header and following lines data. The script expects that there will be data for all dates in the header.

1. Run script:
`python3 trap_activity_gui.py`
2. Add one or more xslx files
3. Click `Process` and select destination of resulting csv file
4. Done