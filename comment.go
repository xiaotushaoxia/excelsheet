package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) AddComment(comment excelize.Comment) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.AddComment(s.Name, comment)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

func (s *Sheet) DeleteComment(cell string) error {
    if s.FirstErr != nil {
       return s.FirstErr
    }
    err := s.File.DeleteComment(s.Name, cell)
    if err != nil {
       s.FirstErr = err
    }
    return err
}

