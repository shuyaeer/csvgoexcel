package main

import (
	"encoding/csv"
	"flag"
	"io"
	"log"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	flag.Parse()
	args := flag.Args()
	file, err := os.Open(args[0])
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	r := csv.NewReader(file)

	line_index := 1
	writer := excelize.NewFile()
	for {
		row, err := r.Read()
		if err == io.EOF {
			break
		}
		create_elsx(row, line_index, writer)
		line_index = line_index + 1
	}
	writer.SaveAs("./Book1.xlsx")

}

func create_elsx(row []string, line_index int, writer *excelize.File) {
	m := map[int]string{0: "A", 1: "B", 2: "C"}
	line_index_str := strconv.Itoa(line_index)
	for i := 0; i < len(row); i++ {
		area := m[i] + line_index_str
		writer.SetCellValue("Sheet1", area, row[i])
	}
}
