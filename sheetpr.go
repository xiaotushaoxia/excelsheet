package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) SetPageMargins(opts *excelize.PageLayoutMarginsOptions) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.SetPageMargins(s.Name, opts)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) SetSheetProps(opts *excelize.SheetPropsOptions) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.SetSheetProps(s.Name, opts)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

