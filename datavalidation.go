package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddDataValidation(dv *excelize.DataValidation) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddDataValidation(s.Name, dv)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) DeleteDataValidation(sqref ...string) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.DeleteDataValidation(s.Name, sqref...)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

