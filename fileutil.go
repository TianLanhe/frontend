package main

import (
	"bytes"
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"strconv"
	"strings"
)

func Byte2Excel(data []byte) (*excelize.File, error) {
	file, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	return file, nil
}

func File2Excel(filename string) (*excelize.File, error) {
	return excelize.OpenFile(filename)
}

func Excel2File(excel *excelize.File, filename string) error {
	return excel.SaveAs(filename)
}

func Excel2Byte(excel *excelize.File) ([]byte, error) {
	buf := bytes.NewBuffer(nil)
	err := Excel2Writer(excel, buf)
	return buf.Bytes(), err
}

func Excel2Writer(excel *excelize.File, w io.Writer) error {
	return excel.Write(w)
}

func Excel2TableStruct(file *excelize.File) (TableStruct, error) {
	var ret TableStruct

	sheets := file.GetSheetList()
	if len(sheets) < 1 {
		return ret, errors.New("没有表Sheet")
	}

	// default read from the first sheet
	rows, err := file.GetRows(sheets[0])
	if err != nil {
		return ret, err
	}

	return parseDataFromArray(rows), nil
}

// given a two-dimension array, the first row is table header, and other row are datas
func parseDataFromArray(rows [][]string) TableStruct {
	var td TableStruct
	maxCol := 0
	for i, row := range rows {
		datas := make([]string, 0, len(row))
		for _, cell := range row {
			datas = append(datas, strings.TrimSpace(cell))
		}

		if i == 0 {
			td.Headers = datas
			continue
		}

		if !allEmptyCells(datas) {
			td.Datas = append(td.Datas, datas)
			if len(datas) > maxCol {
				maxCol = len(datas)
			}
		}
	}
	if maxCol < len(td.Headers) {
		maxCol = len(td.Headers)
	}

	// expand all the data to reach the size of the headers
	// and refresh all row of datas to keep the same size
	for i := range td.Datas {
		for len(td.Datas[i]) < maxCol {
			td.Datas[i] = append(td.Datas[i], "")
		}
	}

	return td
}

// check whether all the elements in slice is empty string
func allEmptyCells(datas []string) bool {
	for _, data := range datas {
		if data != "" {
			return false
		}
	}
	return true
}

func TableStruct2Excel(table TableStruct) *excelize.File {
	sheet := excelize.NewFile()
	for i, header := range table.Headers {
		if header != "" {
			axis := string([]byte{'A' + byte(i), '1'})
			sheet.SetCellStr("Sheet1", axis, header)
		}
	}
	for row, datas := range table.Datas {
		for col, data := range datas {
			if data != "" {
				axis := string([]byte{'A' + byte(col)}) + strconv.Itoa(row+2)
				sheet.SetCellStr("Sheet1", axis, data)
			}
		}
	}
	return sheet
}

func File2TableStruct(filename string) (TableStruct, error) {
	excel, err := File2Excel(filename)
	if err != nil {
		return TableStruct{}, err
	}
	return Excel2TableStruct(excel)
}

func TableStruct2File(table TableStruct, filename string) error {
	excel := TableStruct2Excel(table)
	return Excel2File(excel, filename)
}

func Byte2TableStruct(data []byte) (TableStruct, error) {
	excel, err := Byte2Excel(data)
	if err != nil {
		return TableStruct{}, err
	}
	return Excel2TableStruct(excel)
}

func TableStruct2Byte(table TableStruct) ([]byte, error) {
	excel := TableStruct2Excel(table)
	return Excel2Byte(excel)
}
