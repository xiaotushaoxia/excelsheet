package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) GetCols(opts ...excelize.Options) ([][]string, error) {
	if s.FirstErr != nil {
		return nil, s.FirstErr
	}
	v1, err := s.File.GetCols(s.Name, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) GetColVisible(col string) (bool, error) {
	if s.FirstErr != nil {
		return false, s.FirstErr
	}
	v1, err := s.File.GetColVisible(s.Name, col)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetColVisible(columns string, visible bool) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetColVisible(s.Name, columns, visible)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetColOutlineLevel(col string) (uint8, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetColOutlineLevel(s.Name, col)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetColOutlineLevel(col string, level uint8) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetColOutlineLevel(s.Name, col, level)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetColStyle(columns string, styleID int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetColStyle(s.Name, columns, styleID)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetColWidth(startCol, endCol string, width float64) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetColWidth(s.Name, startCol, endCol, width)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetColStyle(col string) (int, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetColStyle(s.Name, col)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) GetColWidth(col string) (float64, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetColWidth(s.Name, col)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) InsertCols(col string, n int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.InsertCols(s.Name, col, n)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) RemoveCol(col string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.RemoveCol(s.Name, col)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
