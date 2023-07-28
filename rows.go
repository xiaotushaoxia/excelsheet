package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) GetRows(opts ...excelize.Options) ([][]string, error) {
	if s.FirstErr != nil {
		return nil, s.FirstErr
	}
	v1, err := s.File.GetRows(s.Name, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetRowHeight(row int, height float64) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetRowHeight(s.Name, row, height)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetRowHeight(row int) (float64, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetRowHeight(s.Name, row)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetRowVisible(row int, visible bool) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetRowVisible(s.Name, row, visible)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetRowVisible(row int) (bool, error) {
	if s.FirstErr != nil {
		return false, s.FirstErr
	}
	v1, err := s.File.GetRowVisible(s.Name, row)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetRowOutlineLevel(row int, level uint8) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetRowOutlineLevel(s.Name, row, level)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetRowOutlineLevel(row int) (uint8, error) {
	if s.FirstErr != nil {
		return 0, s.FirstErr
	}
	v1, err := s.File.GetRowOutlineLevel(s.Name, row)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) RemoveRow(row int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.RemoveRow(s.Name, row)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) InsertRows(row, n int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.InsertRows(s.Name, row, n)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) DuplicateRow(row int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.DuplicateRow(s.Name, row)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) DuplicateRowTo(row, row2 int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.DuplicateRowTo(s.Name, row, row2)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetRowStyle(start, end, styleID int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetRowStyle(s.Name, start, end, styleID)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
