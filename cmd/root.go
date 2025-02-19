package cmd

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/spf13/cobra"
)

var RootCmd = &cobra.Command{
	Use:   "csvgoexcel",
	Short: "csv converter to excel",
	Run: func(cmd *cobra.Command, args []string) {
		main(args)
	},
}

var columnIndexList map[int]string

func main(args []string) {
	if len(args) == 0 {
		panic("Plese designate file path.")
	}
	filePath := args[0]
	fmt.Println("target file: " + filePath)
	fileExtention := filepath.Ext(filePath)
	if fileExtention == ".csv" {
		csvToExcel(filePath)
	} else if fileExtention == ".xlsx" {
		excelToCsv(filePath)
	}
}

func csvToExcel(filePath string) {
	// able to deal with maximum 52 colums
	columnIndexList = map[int]string{0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F", 6: "G", 7: "H", 8: "I", 9: "J", 10: "K", 11: "L", 12: "M", 13: "M", 14: "O", 15: "p", 16: "Q", 17: "R", 18: "S", 19: "T", 20: "U", 21: "V", 22: "W", 23: "X", 24: "Y", 25: "Z"}
	file, err := os.Open(filePath)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	reader := csv.NewReader(file)
	lineIndex := 1
	writer := excelize.NewFile()
	firstRow := addExtraColumns(reader)
	convertToExcel(firstRow, lineIndex, writer)
	lineIndex = lineIndex + 1
	for {
		row, err := reader.Read()
		if err == io.EOF {
			break
		}
		convertToExcel(row, lineIndex, writer)
		lineIndex = lineIndex + 1
	}
	fileName := createOutputFile(filePath, ".csv")
	writer.SaveAs("./" + fileName)
}

func convertToExcel(row []string, lineIndex int, writer *excelize.File) {
	lineIndexStr := strconv.Itoa(lineIndex)
	for i := 0; i < len(row); i++ {
		area := columnIndexList[i] + lineIndexStr
		writer.SetCellValue("Sheet1", area, row[i])
	}
}

func addExtraColumns(reader *csv.Reader) []string {
	var firstRowCount int
	var row []string
	for i := 0; i < 1; i++ {
		firstRow, err := reader.Read()
		row = firstRow
		if err == io.EOF {
			break
		}
		firstRowCount = len(firstRow)
	}
	if firstRowCount > len(columnIndexList) {
		diff := firstRowCount - len(columnIndexList)
		for i := 0; i < diff; i++ {
			columnIndexList[len(columnIndexList)] = columnIndexList[0] + columnIndexList[i]
		}
	}
	return row
}

func excelToCsv(filePath string) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	outputFileName := strings.Replace(filePath, ".xlsx", "", 1) + ".csv"
	outputFile, err := os.Create(outputFileName)
	if err != nil {
		panic(err)
	}
	defer outputFile.Close()
	w := csv.NewWriter(outputFile)
	rows := file.GetRows("Sheet1")
	for _, record := range rows {
		if err := w.Write(record); err != nil {
			log.Fatalln("faild to write csv", err)
		}
	}
	w.Flush()
	createOutputFile(filePath, ".xlsx")
}

func createOutputFile(filePath string, inputExtention string) string {
	withoutExtention := strings.Replace(filePath, inputExtention, "", 1)
	f := strings.Split(withoutExtention, "/")
	var fileName string
	if inputExtention == ".csv" {
		fileName = f[len(f)-1] + ".xlsx"
	} else {
		fileName = f[len(f)-1] + ".csv"
	}
	fmt.Println(fileName + " is generated!")
	return fileName
}
