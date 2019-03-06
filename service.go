package gsheet

import (
	"context"
	"fmt"
	"io/ioutil"

	"golang.org/x/oauth2/google"
	sheets "google.golang.org/api/sheets/v4"
)

// Service represents a Sheets API service instance.
// Service is the main entry point into using this package.
type Service struct {
	service *sheets.Service
}

// SheetService returns a service to operate on sheet
func SheetService(credentialPath string) (*Service, error) {
	b, err := ioutil.ReadFile(credentialPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.JWTConfigFromJSON(b, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to parse client secret file to config: %v", err)
	}
	client := config.Client(context.TODO())

	s, err := sheets.New(client)
	return &Service{s}, err
}

func (s *Service) FetchSpreadsheet(id string, includeData bool) (spreadsheet *sheets.Spreadsheet, err error) {
	return sheets.NewSpreadsheetsService(s.service).Get(id).IncludeGridData(includeData).Do()
}

/*
// SyncSheet updates sheet
func (s *Service) SyncSheet(sheet *Sheet) (err error) {
	if sheet.newMaxRow > sheet.Properties.GridProperties.RowCount ||
		sheet.newMaxColumn > sheet.Properties.GridProperties.ColumnCount {
		err = s.ExpandSheet(sheet, sheet.newMaxRow, sheet.newMaxColumn)
		if err != nil {
			return
		}
	}
	err = s.syncCells(sheet)
	if err != nil {
		return
	}
	sheet.modifiedCells = []*Cell{}
	sheet.newMaxRow = sheet.Properties.GridProperties.RowCount
	sheet.newMaxColumn = sheet.Properties.GridProperties.ColumnCount
	return
}

// UpdateSpreadsheetTitle update the title of spreadsheet with spreadsheetID
func UpdateSpreadsheetTitle(sheetsService *sheets.Service, spreadsheetID, title string) error {
	req := sheets.Request{
		UpdateSpreadsheetProperties: &sheets.UpdateSpreadsheetPropertiesRequest{
			Fields: "*",
			Properties: &sheets.SpreadsheetProperties{
				Title: title,
			},
		},
	}
	batchRequests := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}
	_, err := sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchRequests).Context(context.Background()).Do()
	return err
}

// CreateSheet create a new sheet for the Spreadsheet with sheetID
func CreateSheet(sheetsService *sheets.Service, spreadsheetID, sheetName string) (int64, error) {
	req := sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title: sheetName,
			},
		},
	}

	rbb := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}

	resp, err := sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, rbb).Context(context.Background()).Do()
	replies := resp.Replies
	if len(replies) == 0 {
		return 0, fmt.Errorf("failed to get ID of the created sheet")
	}
	return replies[0].AddSheet.Properties.SheetId, err
}

func ReadSheet(sheetsService *sheets.Service, spreadsheetID, sheetName, readRange string) ([][]interface{}, error) {
	var emptyData [][]interface{}
	resp, err := sheetsService.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
	if err != nil {
		return emptyData, err
	}

	if len(resp.Values) == 0 {
		return emptyData, fmt.Errorf("no data found")
	}

	return resp.Values, nil
}

// DeleteSheet create a new sheet for the Spreadsheet with sheetID
func DeleteSheet(sheetsService *sheets.Service, spreadsheetID string, sheetID int64) error {
	req := sheets.Request{
		DeleteSheet: &sheets.DeleteSheetRequest{
			SheetId: sheetID,
		},
	}

	rbb := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}

	_, err := sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, rbb).Context(context.Background()).Do()
	return err
}

// updateSheet update the updateRange of sheet with values
func updateSheet(sheetsService *sheets.Service, spreadsheetID, updateRange string, dimension string, values [][]interface{}) error {
	request := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
	}
	request.Data = append(request.Data, &sheets.ValueRange{
		MajorDimension: dimension,
		Range:          updateRange,
		Values:         values,
	})
	_, err := sheetsService.Spreadsheets.Values.BatchUpdate(spreadsheetID, request).Context(context.Background()).Do()
	return err
}

// UpdateSheetByRow update the sheet by row
func UpdateSheetByRow(sheetsService *sheets.Service, spreadsheetID, sheet string, startRow int, startCol byte, values [][]interface{}) error {
	endCol := startCol + byte(len(values[0])-1)
	endRow := startRow + len(values) - 1
	rangeName := fmt.Sprintf("%v!%v%v:%v%v", sheet, string(startCol), startRow, string(endCol), endRow)
	return updateSheet(sheetsService, spreadsheetID, rangeName, "ROWS", values)
}

// UpdateSheetByColumn update the sheet by column
func UpdateSheetByColumn(sheetsService *sheets.Service, spreadsheetID, sheet string, startRow int, startCol byte, values [][]interface{}) error {
	endRow := startRow + len(values[0]) - 1
	// if endCol exceeds 'Z', use 'AA' until 'ZZ'
	// We don't handle column more than 'ZZ' yet
	colCnt := len(values) - 1
	firstByte := byte(colCnt/26-1) + startCol
	secondByte := byte(colCnt%26) + startCol
	endCol := string([]byte{firstByte, secondByte})
	if firstByte == '@' {
		endCol = string(endCol[1])
	}
	rangeName := fmt.Sprintf("%v!%v%v:%v%v", sheet, string(startCol), startRow, string(endCol), endRow)
	return updateSheet(sheetsService, spreadsheetID, rangeName, "COLUMNS", values)
}
*/
