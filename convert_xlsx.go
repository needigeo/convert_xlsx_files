package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"log"

	"github.com/plandem/xlsx"
)

type fileStruct struct {
	fullname  string
	name      string
	sheetname string
}

func main() {
	inputDir := flag.String("inputDir", "null", "Путь к каталогу с исходными файламии. ВАЖНО должен заканчиваться слешем.")
	outputDir := flag.String("outputDir", "null", "Путь к каталогу в который будут сохранены файлы после обработки. ВАЖНО Должен заканчиваться слешем.")
	flag.Parse()

	fileslice := make([]fileStruct, 1)

	if *inputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	if *outputDir == "null" {
		log.Fatal("Не задан входной каталог")
	}

	files, err := ioutil.ReadDir(*inputDir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		fullname := *inputDir + file.Name()

		xl, err := xlsx.Open(fullname)
		if err != nil {
			log.Fatal(err)
		}

		for sheets := xl.Sheets(); sheets.HasNext(); {
			_, sheet := sheets.Next()
			fileslice = append(fileslice, fileStruct{fullname: fullname, name: file.Name(), sheetname: sheet.Name()})
			fmt.Println(fullname, "  ", sheet.Name())
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
		newfullname := *outputDir + num.sheetname + " " + num.name
		fmt.Println(newfullname)
		err = xl.SaveAs(newfullname)
		if err != nil {
			log.Fatal(err)
		}
		fmt.Println(len(dssheets))
		xl.Close()
	}

}
