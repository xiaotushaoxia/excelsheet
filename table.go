package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddTable(table *excelize.Table) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddTable(s.Name, table)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) AutoFilter(rangeRef string, opts []excelize.AutoFilterOptions) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AutoFilter(s.Name, rangeRef, opts)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

