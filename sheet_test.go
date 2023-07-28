package excelsheet

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestExcelizeFile(t *testing.T) {
	file := excelize.NewFile()

	err := file.SetCellValue("testshhe", "A1", 1)
	if err != nil {
		panic(err)
	}
	err = file.SetCellValue("testshhe", "A2", 1)
	if err != nil {
		panic(err)
	}
	err = file.SetCellValue("testshhe", "A3", 1)
	if err != nil {
		panic(err)
	}
	err = file.SetCellValue("testshhe", "A4", 1)
	if err != nil {
		panic(err)
	}
	err = file.SetCellValue("testshhe", "A5", 1)
	if err != nil {
		panic(err)
	}

	buffer, err := file.WriteToBuffer()
	if err != nil {
		panic(err)
	}
	fmt.Println(buffer.Bytes())
}

func TestSheet1(t *testing.T) {
	file := excelize.NewFile()

	sheet, err := New(file, "testshhe")
	if err != nil {
		panic(err)
	}

	err = sheet.SetCellValue("A1", 1)
	if err != nil {
		panic(err)
	}
	err = sheet.SetCellValue("A2", 1)
	if err != nil {
		panic(err)
	}
	err = sheet.SetCellValue("A3", 1)
	if err != nil {
		panic(err)
	}
	err = sheet.SetCellValue("A4", 1)
	if err != nil {
		panic(err)
	}
	err = sheet.SetCellValue("A5", 1)
	if err != nil {
		panic(err)
	}

	buffer, err := file.WriteToBuffer()
	if err != nil {
		panic(err)
	}
	fmt.Println(buffer.Bytes())
}

func TestSheet2(t *testing.T) {
	file := excelize.NewFile()

	sheet, err := New(file, "testshhe")
	if err != nil {
		panic(err)
	}

	sheet.SetCellValue("A1", 1)
	sheet.SetCellValue("A2", 1)
	sheet.SetCellValue("A3", 1)
	sheet.SetCellValue("A4", 1)
	sheet.SetCellValue("A5", 1)

	if err = sheet.FirstErr; err != nil {
		panic(err)
	}

	buffer, err := file.WriteToBuffer()
	if err != nil {
		panic(err)
	}
	fmt.Println(buffer.Bytes())
}

func TestNewStyle(t *testing.T) {
	file := excelize.NewFile()
	fmt.Println(file.NewStyle(&excelize.Style{}))
	fmt.Println(file.NewStyle(&excelize.Style{}))
	fmt.Println(file.NewStyle(&excelize.Style{NumFmt: 1}))
	fmt.Println(file.NewStyle(&excelize.Style{NumFmt: 1}))
}
