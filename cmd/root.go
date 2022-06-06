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

func main(args []string) {
	fileName := args[0]
	fmt.Println(fileName)
	file, err := os.Open(fileName)
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
	fileExt := filepath.Ext(fileName)
	withoutExt := strings.Replace(fileName, fileExt, "", 1)
	f := strings.Split(withoutExt, "/")
	excelFileName := f[len(f)-1] + ".xlsx"
	fmt.Println(excelFileName + " is generated")
	writer.SaveAs("./" + excelFileName)
}

func create_elsx(row []string, line_index int, writer *excelize.File) {
	// need to handle more column,,,
	m := map[int]string{0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F"}
	line_index_str := strconv.Itoa(line_index)
	for i := 0; i < len(row); i++ {
		area := m[i] + line_index_str
		writer.SetCellValue("Sheet1", area, row[i])
	}
}
