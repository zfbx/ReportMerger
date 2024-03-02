package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/term"
)

// "MAIN": {"GL", "GL Name", "Everything", "OPERATIONS", "CGA", "DAF", "TRUST", "PRODUCT", "PASS THROUGH", "FLANIKEN"},
var output = map[string][]interface{}{}
var running = true

func readXls(filename string, col int) {
	if !running {
		return
	}
	dir := "./"
	files, err := os.ReadDir(dir)
	if err != nil {
		log.Fatal(err)
	}
	for _, file := range files {
		if strings.HasPrefix(file.Name(), filename) {
			dir = dir + file.Name()
			fmt.Println("Located", filename, "file: \"", file.Name(), "\"")
		}
	}

	if dir == "./" {
		if filename == "BASE" {
			fmt.Println("To use this script properly you need to have these 8 xlsx files in the same folder as this script\n\nBASE (Everything - REQUIRED)\nCGA\nDAF\nFLANIKEN (trust)\nOPERATIONS\nPASS (passthrough)\nPRODUCT\nTRUST (without flaniken)\n\nAdd them and re-run the script :)")
			running = false
			return
		}
		fmt.Println("CAN'T FIND FILE FOR", filename)
		return
	}

	f, err := excelize.OpenFile(dir)
	if err != nil {
		panic(err)
	}
	defer closeFile(f)

	rows, err := f.GetRows("Report")
	if err != nil {
		panic(err)
	}
	for id, row := range rows {
		if id == 0 {
			continue
		}
		if _, err = strconv.Atoi(row[0]); err != nil {
			continue
		}

		var err error
		dr := 0.0
		cr := 0.0
		if len(row) > 6 && len(row[6]) > 1 {
			if dr, err = strconv.ParseFloat(strings.ReplaceAll(row[6][1:], ",", ""), 64); err != nil {
				fmt.Println("Error converting ammount: ", row[6])
			}
		}
		if len(row) > 7 {
			if cr, err = strconv.ParseFloat(strings.ReplaceAll(row[7][1:], ",", ""), 64); err != nil {
				fmt.Println("Error converting ammount: ", row[6])
			}
		}

		if col == 0 { // base
			output[row[0]] = []interface{}{row[0], row[1], dr - cr, 0, 0, 0, 0, 0, 0, 0}
		} else { // other
			output[row[0]][col] = dr - cr
		}
	}
}

func writeXls() {
	f := excelize.NewFile()
	defer closeFile(f)

	sheet, err := f.NewSheet("Report")
	if err != nil {
		fmt.Println(err)
		return
	}

	f.SetCellValue("Report", "A1", "GL")
	f.SetCellValue("Report", "B1", "GL NAME")
	f.SetCellValue("Report", "C1", "EVERYTHING")
	f.SetCellValue("Report", "D1", "OPERATIONS")
	f.SetCellValue("Report", "E1", "CGA")
	f.SetCellValue("Report", "F1", "DAF")
	f.SetCellValue("Report", "G1", "TRUST")
	f.SetCellValue("Report", "H1", "PRODUCT")
	f.SetCellValue("Report", "I1", "PASS THROUGH")
	f.SetCellValue("Report", "J1", "FLANIKEN")

	style, err := f.NewStyle(&excelize.Style{NumFmt: 4})
	if err != nil {
		fmt.Println("Error creating style:", err)
		return
	}

	index := 2
	for _, row := range output {
		for ci, col := range row {
			colName, _ := excelize.ColumnNumberToName(ci + 1)
			cell := colName + strconv.Itoa(index)
			f.SetCellValue("Report", cell, col)

			if ci > 1 {
				f.SetCellStyle("Report", cell, cell, style)
			}
		}
		index++
	}
	f.SetActiveSheet(sheet)
	savename := "Merged Reports.xlsx"
	if err := f.SaveAs(savename); err != nil {
		fmt.Println(err)
	}
	fmt.Println("All done! Saved as", savename)
}

func closeFile(f *excelize.File) {
	err := f.Close()
	if err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}
}

func main() {
	fmt.Println("Report Merger - Made with ♥")

	readXls("BASE", 0)
	readXls("OPERATIONS", 3)
	readXls("CGA", 4)
	readXls("DAF", 5)
	readXls("TRUST", 6)
	readXls("PRODUCT", 7)
	readXls("PASS", 8)
	readXls("FLANIKEN", 9)

	if running {
		writeXls()
	}

	if term.IsTerminal(int(os.Stdout.Fd())) {
		fmt.Println("Press any key to close..")
		var input string
		fmt.Scanln(&input)
	}
}
