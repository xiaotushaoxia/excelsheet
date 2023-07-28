package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) GetCellStyle(cell string) (int, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetCellStyle(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetCellStyle(hCell, vCell string, styleID int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellStyle(s.Name, hCell, vCell, styleID)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetConditionalFormat(rangeRef string, opts []excelize.ConditionalFormatOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetConditionalFormat(s.Name, rangeRef, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) UnsetConditionalFormat(rangeRef string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.UnsetConditionalFormat(s.Name, rangeRef)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
