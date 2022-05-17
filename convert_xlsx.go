package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"strings"

	"github.com/plandem/xlsx"
	"github.com/xuri/excelize/v2"
)

type fileStruct struct {
	fullname  string
	name      string
	sheetname string
	maxcol    int
	maxrow    int
}

func main() {
	inputDir := flag.String("inputDir", "null", "Путь к каталогу с исходными файламии. ВАЖНО должен заканчиваться слешем.")
	outputDir := flag.String("outputDir", "null", "Путь к каталогу в который будут сохранены файлы после обработки. ВАЖНО Должен заканчиваться слешем.")
	flag.Parse()
	convwithexcelize(*inputDir, *outputDir)
}

func convwithxlsx(inputDir string, outputDir string) {
	fileslice := make([]fileStruct, 1)

	if inputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	if outputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	files, err := ioutil.ReadDir(inputDir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		fullname := inputDir + file.Name()

		xl, err := xlsx.Open(fullname)
		if err != nil {
			log.Fatal(err)
		}

		for sheets := xl.Sheets(); sheets.HasNext(); {
			_, sheet := sheets.Next()
			if sheet.Name() != "Классификация" {
				fileslice = append(fileslice, fileStruct{fullname: fullname, name: file.Name(), sheetname: sheet.Name()})
				fmt.Println(fullname, "  ", sheet.Name())
			}
		}

		xl.Close()
		//fmt.Println(fullname, file.Name(), file.IsDir())
	}

	for _, num := range fileslice[1:] {
		dssheets := []int{99999}
		i := 0
		xl, err := xlsx.Open(num.fullname)
		if err != nil {
			log.Fatal(err)
		}

		for sheets := xl.Sheets(); sheets.HasNext(); {
			_, sheet := sheets.Next()
			if sheet.Name() != num.sheetname {
				dssheets = append(dssheets, i)
			}
			i = i + 1
		}
		for _, index := range dssheets[1:] {
			fmt.Println(index)
			xl.DeleteSheet(index)
		}
		newfullname := outputDir + num.sheetname + " " + num.name
		fmt.Println(newfullname)
		err = xl.SaveAs(newfullname)
		if err != nil {
			log.Fatal(err)
		}
		fmt.Println(len(dssheets))
		err = xl.Close()
		if err != nil {
			log.Fatal(err)
		}
	}
}

func convwithexcelize(inputDir string, outputDir string) {
	fileslice := make([]fileStruct, 1)

	if inputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	if outputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	files, err := ioutil.ReadDir(inputDir)
	lerr(err)

	for _, file := range files {
		fullname := inputDir + file.Name()
		contain := strings.Contains(fullname, "xlsx")
		if contain != true {
			continue
		}
		xl, err := excelize.OpenFile(fullname)
		lerr(err)

		for _, name := range xl.GetSheetMap() {
			cols, _ := xl.GetCols(name)
			rows, _ := xl.GetRows(name)

			maxcol := len(cols)
			maxrow := len(rows)
			fileslice = append(fileslice, fileStruct{fullname: fullname, name: file.Name(), sheetname: name, maxcol: maxcol, maxrow: maxrow})

		}

		xl.Close()
		//fmt.Println(fullname, file.Name(), file.IsDir())
	}

	for _, num := range fileslice[1:] {
		xl, err := excelize.OpenFile(num.fullname)
		lerr(err)

		newxl := excelize.NewFile()
		for row := 1; row <= num.maxrow; row++ {
			for col := 1; col <= num.maxcol; col++ {
				if col == 0 {
					col = 1
				}
				axis, err := excelize.CoordinatesToCellName(col, row)
				lerr(err)
				value, err := xl.GetCellValue(num.sheetname, axis)
				lerr(err)

				newxl.SetCellValue("Sheet1", axis, value)

			}

		}

		newfullname := outputDir + num.sheetname + " " + num.name

		err = newxl.SaveAs(newfullname)
		lerr(err)
		err = xl.Close()
		lerr(err)
		err = newxl.Close()
		lerr(err)
	}
}

func lerr(err error) {
	if err != nil {
		log.Fatal(err)
	}
}
