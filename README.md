# JNW Engineering Job Order

Standalone mobile Job Order form for JNW Engineering.

## Start

Double-click `Start Job Order Server.bat` on the office PC.

Technicians open:

```text
http://192.168.10.38:4322/
```

## Reports

When technicians press `Submit to OneDrive`, reports are saved in:

```text
Reports/
```

Each submission creates:

- an `.html` report for viewing/printing
- a `.json` data file
- a row in `Job_Order_Log.csv`

Do not upload the `Reports` folder to GitHub because it contains real job records.
