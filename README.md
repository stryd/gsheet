Package `gsheet` provides easy-to-use access to the Google Sheets API for reading and updating Google Spreadsheets.

## Usage

First you need **service** to start using this package. You also need service account key from Google Developer Console to init the **service**.

```go
service, err := gsheet.SheetService("client_secret.json")
checkError(err)
```

### Fetching a spreadsheet with all the contents of the sheets included

```go
spreadsheetID := "1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4"
sheet, err := service.FetchSpreadsheet(spreadsheetID, includeData)
checkError(err)
```
