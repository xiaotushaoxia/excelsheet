package excelsheet

func (s *Sheet) MergeCell(hCell, vCell string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.MergeCell(s.Name, hCell, vCell)
	if err != nil {
		s.FirstErr = err
	}
	return err
}

func (s *Sheet) UnmergeCell(hCell, vCell string) error {
	if s.FirstErr != nil {
		return s.FirstErr
	}
	err := s.File.UnmergeCell(s.Name, hCell, vCell)
	if err != nil {
		s.FirstErr = err
	}
	return err
}
