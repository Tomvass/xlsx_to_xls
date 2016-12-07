package main

import (
	"archive/zip"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"io"

	"github.com/aswjh/excel"
)

func convertFile(inputFile, outputFile string) {
	options := excel.Option{"Visible": false, "DisplayAlerts": false}
	xl, err := excel.Open(inputFile, options)
	if err != nil {
		log.Fatal(err)
	}
	defer xl.Quit()
	xl.SaveAs(outputFile, 56)
	log.Printf("Successfully converted: %s.\n", outputFile)
}

func convertZipfile(zipfile, outDir string) {
	r, err := zip.OpenReader(zipfile)
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()
	for _, f := range r.File {
		if strings.HasSuffix(f.Name, ".xlsx") {
			tmpfile, _ := ioutil.TempFile(outDir, f.Name)
			defer os.Remove(tmpfile.Name())
			dt, _ := f.Open()
			io.Copy(tmpfile, dt)
			outName := strings.Split(tmpfile.Name(), ".")[0] + ".xls"
			dt.Close()
			tmpfile.Close()
			convertFile(tmpfile.Name(), outName)
		}
	}
}

func main() {
	lookDir, _ := filepath.Abs(filepath.Dir(os.Args[0]))
	files, _ := ioutil.ReadDir(lookDir)
	log.Println(lookDir)
	for _, f := range files {
		if !f.IsDir() {
			if strings.HasSuffix(f.Name(), ".xlsx") {
				inName := lookDir + "\\" + f.Name()
				outName := lookDir + "\\" + strings.Split(f.Name(), ".")[0] + ".xls"
				convertFile(inName, outName)
			}
			if strings.HasSuffix(f.Name(), ".zip") {
				zipfile := lookDir + "\\" + f.Name()
				convertZipfile(zipfile, lookDir)
			}
		}
	}
}
