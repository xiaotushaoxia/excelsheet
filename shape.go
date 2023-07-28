package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddShape(cell string, opts *excelize.Shape) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddShape(s.Name, cell, opts)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

