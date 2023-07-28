package excelsheet

import "github.com/xuri/excelize/v2"

func (s *Sheet) SetSheetView(viewIndex int, opts *excelize.ViewOptions) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.SetSheetView(s.Name, viewIndex, opts)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) GetSheetView(viewIndex int) (excelize.ViewOptions, error) {
	if s.FirstErr != nil {
		return excelize.ViewOptions{
			DefaultGridColor:  nv(true),
			ShowFormulas:      nv(true),
			ShowGridLines:     nv(true),
			ShowRowColHeaders: nv(true),
			ShowRuler:         nv(true),
			ShowZeros:         nv(true),
			View:              nv("normal"),
			ZoomScale:         nv(100.),
		}, s.FirstErr
	}
	v1, err := s.File.GetSheetView(s.Name, viewIndex)
	if err != nil {
		s.FirstErr = err
	}
	return v1, err
}

func nv[T any](t T) *T {
	return &t
}
