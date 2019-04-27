package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"go_excel/library"
	"os"
	"strconv"
	"time"
)

type Salary struct {
	Data map[string][]Type `json:"data"`
}
type OneType struct {
	Type  string
	Value int
}
type Type struct {
	EmployeeId int `json:"employee_id"`
	Value      int `json:"value"`
}
type CellHeader struct {
	RowIndex int
	Axis     string
	CellName string
	Letter   string
}
type DataExcel struct {
	NumberColumn int    `json:"numberColumn"`
	Lable        string `json:"lable"`
	Letter       string `json:"letter"`
	Value        string `json:"value"`
}
type AxisCell struct {
	Int    int
	Letter string
	Axis   string
	Key    string
	Value  string
}

type ParseExcel struct {
	Data []DataExcel `json:"data"`
}

func main() {
	jsonFile, err := os.Open("test.json")
	if err != nil {
		fmt.Println("Error opening test file\n", err.Error())
		return
	}
	jsonParser := json.NewDecoder(jsonFile)
	var salary Salary
	if err = jsonParser.Decode(&salary); err != nil {
		fmt.Println("Error while reading test file.\n", err.Error())
		return
	}
	filePath := "Cau_Hinh_Bang_Luong_" + time.Now().Format("2016-07-06") + ".xlsx"
	WriteToNewFile(filePath, "Sheet1", salary)
}

func WriteToNewFile(filePath, sheet string, salary Salary) {
	xlsx := excelize.NewFile()
	index := xlsx.GetSheetIndex(sheet)
	xlsx.SetActiveSheet(index)
	listExcel, listCellQuery, parseExcel := ReadExcel(salary)
	res2B, _ := json.Marshal(parseExcel)
	fmt.Println(string(res2B))
	for _, value := range listExcel {
		if value.RowIndex == 0 {
			_ = xlsx.SetCellStr(sheet, value.Axis, value.CellName)
			_ = xlsx.SetColWidth(sheet, "C", "BT", 20)
		}
	}
	listEmployeeId := []int{}
	listType  := map[string][]Type{}
	for key , value := range salary.Data {
		listTypeValue := []Type{}
		for _, valueKey := range value {
			typeValue := Type{}
			listEmployeeId = append(listEmployeeId, valueKey.EmployeeId)
			typeValue.Value = valueKey.Value
			typeValue.EmployeeId = valueKey.EmployeeId
			listTypeValue = append(listTypeValue, typeValue)
		}
		listType[key]= listTypeValue
	}
	listEmployeeId = library.RemoveDuplicates(listEmployeeId)
	n := 2
	for _, value := range listEmployeeId {
		WriteExcelValue(xlsx, "B"+strconv.Itoa(n), strconv.Itoa(value), sheet)
		for _ , v :=range listCellQuery{
			for _ , m :=range listType[v.Key]{
				if m.EmployeeId == value {
					WriteExcelValue(xlsx, v.Letter+strconv.Itoa(n), strconv.Itoa(m.Value), sheet)
				}
			}

		}
		n += 1
	}
	err := xlsx.SaveAs(filePath)
	if err != nil {
		fmt.Println(err)
	}
}
func WriteExcelValue(xlsx *excelize.File, axis, value, sheet string) string {
	_ = xlsx.SetCellStr(sheet, axis, value)
	return value
}
func ReadExcel(salary Salary) ([]CellHeader, []AxisCell, ParseExcel) {
	xlsx, err := excelize.OpenFile("Cau_Hinh_Bang_Luong.xlsx")
	if err != nil {
		fmt.Println("Read Error", err.Error())
	}
	//b2, _ := xlsx.GetCellFormula("Sheet1", "B2")
	//fmt.Println(b2)
	//get tat ca cac cell
	rows, _ := xlsx.GetRows("Sheet1")
	// lap lai tung dong ，
	var listCellHeader []CellHeader
	var listAxisCellQuery []AxisCell
	//filePath := "Cau_Hinh_Bang_Luong" + time.Now().Format("2016-07-06") + ".xlsx"
	for rowIndex, row := range rows {
		for cellIndex, cell := range row {
			letter, axis, _ := library.PositionToAxis(rowIndex, cellIndex)
			singleCell := CellHeader{}
			if rowIndex == 0 {
				singleCell.Axis = axis
				singleCell.RowIndex = rowIndex
				singleCell.CellName = cell
				singleCell.Letter = letter
				listCellHeader = append(listCellHeader, singleCell)
			}
			for key, value := range salary.Data {
				for range value {
					if cell == key {
						axiscel := AxisCell{}
						letter, axisCell, num := library.PositionToAxis(rowIndex, cellIndex)
						axiscel.Int = num
						axiscel.Letter = letter
						axiscel.Axis = axisCell
						axiscel.Key = key
						listAxisCellQuery = append(listAxisCellQuery, axiscel)
					}
				}

			}
		}
	}
	parseExcel := ParseExcel{}
	for rowIndex, row := range rows {
		if rowIndex == 1 {
			break
		}
		for cellIndex, cell := range row {
			parseExcelSingle := DataExcel{}
			letter, Axis, _ := library.PositionToAxis(rowIndex, cellIndex)
			valueFomula, _ := xlsx.GetCellFormula("Sheet1", Axis)
			fmt.Print(valueFomula)
			if valueFomula != "" {
				parseExcelSingle.Value = valueFomula
			} else {
				parseExcelSingle.Value = cell
			}
			parseExcelSingle.NumberColumn = cellIndex
			parseExcelSingle.Lable = listCellHeader[cellIndex].CellName
			parseExcelSingle.Letter = letter
			parseExcel.Data = append(parseExcel.Data, parseExcelSingle)
		}
	}

	return listCellHeader, listAxisCellQuery, parseExcel
}
