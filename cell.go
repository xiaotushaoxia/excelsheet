package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) GetCellValue(cell string, opts ...excelize.Options) (string, error) {
	if s.FirstErr != nil {
		return "", s.FirstErr
	}
	v1, err := s.File.GetCellValue(s.Name, cell, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) GetCellType(cell string) (excelize.CellType, error) {
	if s.FirstErr != nil {
		return excelize.CellTypeUnset, s.FirstErr
	}
	v1, err := s.File.GetCellType(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetCellValue(cell string, value interface{}) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellValue(s.Name, cell, value)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetCellInt(cell string, value int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellInt(s.Name, cell, value)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetCellBool(cell string, value bool) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellBool(s.Name, cell, value)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetCellFloat(cell string, value float64, precision, bitSize int) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellFloat(s.Name, cell, value, precision, bitSize)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetCellStr(cell, value string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellStr(s.Name, cell, value)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetCellDefault(cell, value string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellDefault(s.Name, cell, value)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetCellFormula(cell string) (string, error) {
	if s.FirstErr != nil {
		return "", s.FirstErr
	}
	v1, err := s.File.GetCellFormula(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetCellFormula(cell, formula string, opts ...excelize.FormulaOpts) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellFormula(s.Name, cell, formula, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetCellHyperLink(cell string) (bool, string, error) {
	if s.FirstErr != nil {
		return false, "", s.FirstErr
	}
	v1, v2, err := s.File.GetCellHyperLink(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, v2, err
}

func (s *Sheet) SetCellHyperLink(cell, link, linkType string, opts ...excelize.HyperlinkOpts) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellHyperLink(s.Name, cell, link, linkType, opts...)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetCellRichText(cell string) (runs []excelize.RichTextRun, err error) {
	if s.FirstErr != nil {
		return nil, s.FirstErr
	}
	v1, err := s.File.GetCellRichText(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetCellRichText(cell string, runs []excelize.RichTextRun) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetCellRichText(s.Name, cell, runs)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetSheetRow(cell string, slice interface{}) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetRow(s.Name, cell, slice)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetSheetCol(cell string, slice interface{}) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetCol(s.Name, cell, slice)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
