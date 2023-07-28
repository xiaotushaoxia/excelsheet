package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddChart(cell string, chart *excelize.Chart, combo ...*excelize.Chart) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddChart(s.Name, cell, chart, combo...)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) AddChartSheet(chart *excelize.Chart, combo ...*excelize.Chart) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddChartSheet(s.Name, chart, combo...)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) DeleteChart(cell string) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.DeleteChart(s.Name, cell)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

