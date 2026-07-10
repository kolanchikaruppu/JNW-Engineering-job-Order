# Shared Job Order Number Setup

This form needs one central Microsoft system to issue job order numbers. Do not use a local browser counter when many technicians submit jobs at the same time.

## 1. Create A Microsoft List

Create a SharePoint or Microsoft List named `JNW Job Orders`.

Recommended columns:

- `Title` - single line text
- `Company` - single line text
- `Date` - date
- `RequestedBy` - single line text
- `Complaint` - multiple lines text
- `ActionTaken` - multiple lines text
- `LabourDescription` - single line text
- `LabourMan` - number
- `LabourHours` - number
- `Remarks` - multiple lines text
- `PartsJson` - multiple lines text

The built-in list `ID` column is the shared sequence number. It is created by SharePoint automatically and is safe when many technicians submit at once.

## 2. Create A Power Automate Flow

Create an automated cloud flow with this trigger:

`When an HTTP request is received`

Use this sample JSON body:

```json
{
  "jobOrder": "Pending",
  "company": "Customer name",
  "date": "2026-07-10",
  "requestedBy": "Requester",
  "complaint": "Complaint text",
  "actionTaken": "Action taken text",
  "labourDescription": "Labour",
  "labourMan": "1",
  "labourHours": "2",
  "remarks": "Remarks",
  "parts": [
    {
      "serial": "1",
      "description": "Part",
      "charge": "2"
    }
  ]
}
```

Add a `Create item` action for the `JNW Job Orders` list.

Map fields from the HTTP body. For `Title`, use:

`Job Order`

For `PartsJson`, use:

`string(triggerBody()?['parts'])`

## 3. Return The Official Job Number

Add a `Response` action at the end.

Status code:

`200`

Headers:

```text
Access-Control-Allow-Origin: *
Content-Type: application/json
```

Body:

```json
{
  "jobOrder": "@{formatNumber(outputs('Create_item')?['body/ID'], '00')}"
}
```

## 4. Connect The Form

Copy the HTTP POST URL from the Power Automate trigger.

Open `index.html`, find this line:

```js
const submitEndpoint = '';
```

Paste the URL between the quotes.

## 5. PDF Saving

For automatic PDF saving into SharePoint or OneDrive, extend the same Power Automate flow after `Create item`.

Recommended file name:

```text
JNW Job Order @{formatNumber(outputs('Create_item')?['body/ID'], '00')}.pdf
```

The browser form can request the central job number and open printing, but the browser cannot silently write a PDF into a PC OneDrive folder without a save dialog.
