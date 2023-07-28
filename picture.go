package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddPicture(cell, name string, opts *excelize.GraphicOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.AddPicture(s.Name, cell, name, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) AddPictureFromBytes(cell string, pic *excelize.Picture) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.AddPictureFromBytes(s.Name, cell, pic)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetPictures(cell string) ([]excelize.Picture, error) {
	if s.FirstErr != nil {
		return nil, s.FirstErr
	}
	v1, err := s.File.GetPictures(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func (s *Sheet) DeletePicture(cell string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.DeletePicture(s.Name, cell)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
