package excelsheet

import "github.com/xuri/excelize/v2"

type Sheet struct {
	File     *excelize.File
	Name     string
	Index    int
	FirstErr error
}

func New(file *excelize.File, name string) (*Sheet, error) {
	var idx int
	var err error
	if _, ok := file.Sheet.Load(name); ok {
		idx, err = file.NewSheet(name)
	} else {
		idx, err = file.GetSheetIndex(name)
	}
	if err != nil {
		return nil, err
	}
	return &Sheet{
		File:  file,
		Name:  name,
		Index: idx,
	}, nil
}
func (s *Sheet) SetSheetBackground(picture string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetBackground(s.Name, picture)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetSheetBackgroundFromBytes(extension string, picture []byte) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetBackgroundFromBytes(s.Name, extension, picture)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetSheetVisible(visible bool, veryHidden ...bool) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetVisible(s.Name, visible, veryHidden...)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetPanes(panes *excelize.Panes) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetPanes(s.Name, panes)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SearchSheet(value string, reg ...bool) ([]string, error) {
	if s.FirstErr != nil {
		return nil, s.FirstErr
	}
	v1, err := s.File.SearchSheet(s.Name, value, reg...)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) SetHeaderFooter(opts *excelize.HeaderFooterOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetHeaderFooter(s.Name, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) ProtectSheet(opts *excelize.SheetProtectionOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.ProtectSheet(s.Name, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) UnprotectSheet(password ...string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.UnprotectSheet(s.Name, password...)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetPageLayout(opts *excelize.PageLayoutOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetPageLayout(s.Name, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) InsertPageBreak(cell string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.InsertPageBreak(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) RemovePageBreak(cell string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.RemovePageBreak(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) SetSheetDimension(rangeRef string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetDimension(s.Name, rangeRef)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
