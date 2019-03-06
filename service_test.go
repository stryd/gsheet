package gsheet

import (
	"testing"
)

func TestSheetIntegration(t *testing.T) {
	filepath := "/Users/kunli/Work/golang/src/github.com/stryd/cloud/backend/keys/internalservice-f083df03884d.json"
	service, err := SheetService(filepath)
	if err != nil {
		t.Errorf("Failed init sheet service: %v", err)
		return
	}
	spreadsheetID := "1qTj2X22bdf3BDxE_BdT1UjnT9-H2d072zHhLutbDyvs"
	ss, err := service.FetchSpreadsheet(spreadsheetID, true)
	if err != nil {
		t.Errorf("Failed fetch spread service: %v", err)
		return
	}
	t.Logf("There are %v sheets in this spreadsheet", len(ss.Sheets))
	for k, s := range ss.Sheets {
		if k == 0 {
			continue
		}
		t.Logf("sheet %v", s.Properties.Title)
		data := s.Data
		for i, d := range data {
			t.Logf("grid %v", i)
			t.Logf("There are %v user emails", len(d.RowData))
			for _, r := range d.RowData {
				for iii, cv := range r.Values {
					if cv.EffectiveValue != nil {
						t.Logf("column %v: %v", iii, cv.EffectiveValue.StringValue)
					}
				}
			}
		}
	}
}
