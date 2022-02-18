package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

//根据全量Excel文件中内容获取相应数据
func fetchDataByExcel() {
	//fmt.Println(filepath + "全量数据文件.xlsx")
	f, err := excelize.OpenFile(filepath + "全量数据文件.xlsx")

	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	tplf, err := excelize.OpenFile(filepath + "模板文件.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	//tplsrows, err := tplf.GetRows("Sheet1")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}

	var fundName string = "默认基金名称"
	var iCol int = 2
	var iAllCol int = 0
	for _, row := range rows {
		//fmt.Println(row[6])
		if getFileName(row[6]) == "" {
			continue
		}

		if fundName == "默认基金名称" {
			fundName = row[6]
		}

		if row[6] != fundName {
			iAllCol++
			iCol = 2
			//fmt.Println("基金发生变更")
			fmt.Println(getFileName(fundName))

			tplf.SaveAs("生成文件/" + getFileName(fundName))
			fundName = row[6]
			//time.Sleep(time.Duration(2) * time.Second)
			//重新打开模板文件，不然数据有问题
			//tplf.Close()
			tplf, err = excelize.OpenFile(filepath + "模板文件.xlsx")
			if err != nil {
				fmt.Println(err)
				return
			}
			if err != nil {
				fmt.Println(err)
				return
			}

		}

		if iCol > 2 {
			err = tplf.DuplicateRow("Sheet1", iCol-1)
			if err != nil {
				fmt.Println(err)
				return
			}
		}

		tplf.SetSheetRow("Sheet1", "A"+strconv.Itoa(iCol), &[]interface{}{row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13],
			row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29]})
		iCol++
		iAllCol++

	}
	//最后一个基金，需要调用重新生成
	fmt.Println(getFileName(fundName))
	tplf.SaveAs("生成文件/" + getFileName(fundName))
	//time.Sleep(time.Duration(2) * time.Second)
	//重新打开模板文件，不然数据有问题
	//tplf.Close()
	tplf, err = excelize.OpenFile(filepath + "模板文件.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("共导出文件", iAllCol, "个")
}

//根据全量数据中的基金名称获取导出文件名称，未获取到则返回空文件
func getFileName(FundFullName string) string {
	f, err := excelize.OpenFile(filepath + "全量数据文件.xlsx")
	if err != nil {
		fmt.Println(err)
		return "err"
	}
	rows, err := f.GetRows("Sheet2")
	if err != nil {
		fmt.Println(err)
		return "err"
	}
	for _, row := range rows {
		if row[1] == FundFullName {
			return row[0] + row[2] + ".xlsx"
		}
		//fmt.Println()
	}
	return ""
}
