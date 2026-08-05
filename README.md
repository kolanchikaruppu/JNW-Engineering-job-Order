# JNW Engineering Job Order

Standalone mobile Job Order form for JNW Engineering.

## Link

Technicians open the Azure web app link.

## Reports

When technicians press `Submit`, the completed report is saved as a PDF in company OneDrive/SharePoint:

```text
Documents / Job Order Reports / 2026
```

The year folder is created automatically from the job order date. For example,
2027 reports go to `Documents / Job Order Reports / 2027`.

PDF file name format:

```text
JO XX _ Customer Name _ Date.pdf
```

Example:

```text
JO 03 _ JNW _ 31-07-2026.pdf
```

## Notes

The Azure web app needs SharePoint permission before submissions can upload PDFs.
To save into a specific OneDrive instead, set the Azure app setting
`GRAPH_DRIVE_ID` to that OneDrive drive ID.
