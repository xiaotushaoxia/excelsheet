package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) CalcCellValue(cell string, opts ...excelize.Options) (result string, err error) {
	if s.FirstErr != nil {
		return "", s.FirstErr
	}
	v1, err := s.File.CalcCellValue(s.Name, cell, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}
