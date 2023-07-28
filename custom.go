package excelsheet

import (
	"github.com/xuri/excelize/v2"
)

func (s *Sheet) SetRowsHeight(startRow, endRow int, height float64) error {
	for i := startRow; i < endRow+1; i++ {
		s.SetRowHeight(i, height)
	}
	return s.FirstErr
}

func (s *Sheet) AddCellStyle(hCell, vCell string, style *excelize.Style) error {
	styleID, err := s.newStyle(style)
	if err != nil {
		return err
	}
	return s.SetCellStyle(hCell, vCell, styleID)
}

func (s *Sheet) AddRowStyle(start, end int, style *excelize.Style) error {
	styleID, err := s.newStyle(style)
	if err != nil {
		return err
	}
	return s.SetRowStyle(start, end, styleID)
}

// AddColStyle err = s.AddColStyle("C:F", style)
func (s *Sheet) AddColStyle(columns string, style *excelize.Style) error {
	styleID, err := s.newStyle(style)
	if err != nil {
		return err
	}
	return s.SetColStyle(columns, styleID)
}

func (s *Sheet) newStyle(style *excelize.Style) (int, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}

	styleID, err := s.File.NewStyle(style)
	if err != nil {
		s.FirstErr = err
		return 0, err
	}
	return styleID, nil
}
