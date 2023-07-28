package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddSparkline(opts *excelize.SparklineOptions) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddSparkline(s.Name, opts)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

