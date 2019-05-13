package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

type record struct {
	number     int
	branchNo   string
	branchName string
	address    string
	phoneNo    string
	quantity   int
}

func main() {
	//f, err := excelize.OpenFile("./info.xlsx")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//// Get value from cell by given worksheet name and axis.
	//cell, err := f.GetCellValue("record", "B2")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//fmt.Println(cell)
	fmt.Println("Generate print paper program is starting")
	readRecord()


}

func readRecord(){
	fmt.Println("read input from source")
	f, err := excelize.OpenFile("./source_record/info.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Get all the rows in the record sheet.
	rows, err := f.GetRows("record")
	recordOfInfo := []record{}
	for _, row := range rows {
		orderOfCell := 0
		tempRecord := record{}
		for _, colCell := range row {
			orderOfCell = orderOfCell + 1
			if orderOfCell == 1 {
				if _, err := strconv.Atoi(colCell); err != nil {
					//check if the first column is not a number
					//fmt.Printf("%q is not a number.\n", colCell)
					break
				}
				tempRecord.number, err = strconv.Atoi(colCell)
				if err != nil {
					fmt.Println(colCell, " is not a number please check")
				}
			} else if orderOfCell == 2 {
				tempRecord.branchNo = colCell
			} else if orderOfCell == 3 {
				tempRecord.branchName = colCell
			} else if orderOfCell == 4 {
				tempRecord.address = colCell
			} else if orderOfCell == 5 {
				tempRecord.phoneNo = colCell
			} else if orderOfCell == 6 {
				tempRecord.quantity, err = strconv.Atoi(colCell)
				if err != nil {
					fmt.Println(colCell, " is not a number please check")
				}
			}
			//fmt.Print(colCell, "\t")
		}
		if tempRecord.number != 0 {
			recordOfInfo = append(recordOfInfo, tempRecord)
		}

		//fmt.Println()
	}

	fmt.Println("recordOfInfo is ", recordOfInfo)
}
