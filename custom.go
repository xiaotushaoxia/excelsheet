package excelsheet

func (s *Sheet) SetRowsHeight(startRow, endRow int, height float64) error {
	for i := startRow; i < endRow+1; i++ {
		s.SetRowHeight(i, height)
	}
	return s.FirstErr
}
