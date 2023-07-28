## 对excelize中sheet的封装
1. 把sheet当成操作对象，可以避免在文件操作中频繁带sheet参数
2. 封装了操作错误，可以不在每次操作后都判断err，在最后判断即可
3. 发现excelize中行高需要SetRowHeight一行行设置（？暂时没找到），增加了批量设置行高的函数SetRowsHeight

## Example

```go usage
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
```