/*
  load data from json file and translate the data to a xlsx file

json format
 {
	"data":{
		"sheet_name":{
			"col_format":[{"col":"A:A", "format":{"width":15}},{"col":"B:B", "format":{"width":15}}],
			"rows":[
				["value":"name", "format":{"align":"center","valign":"middle", "bg_color":"#cccccc"}],
				["value":10, "format":{"align":"center","valign":"middle", "bg_color":"#cccccc"}]
			]
		}
	}

*/

package main

import (
	"flag"
	"fmt"
	"os"
	//"io"
	"encoding/json"
	"io/ioutil"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/bitly/go-simplejson"
	"github.com/json-iterator/go"
	"github.com/koron/go-dproxy"
	"github.com/tealeg/xlsx"

	"excel/common"
)

var jsonFile string
var debug, jsonMethod int

func loadJson() string {
	//os.Exit(0)
	str, err := ioutil.ReadFile(jsonFile)
	if err == nil {
		return string(str)
	} else {
		fmt.Println(err)
		return ""
	}
}

func _goProxy() {
	var v interface{}
	jsonStr := loadJson()
	json.Unmarshal([]byte(jsonStr), &v)
	sheet := dproxy.Pointer(v, "/data/s1") //.String() //M("data").M("s1").A(0).A(0).M("value").String()
	row := sheet.A(0)
	col := row.A(0)
	fmt.Println(col.M("value").String())
	fmt.Println(col.M("format").M("bold").Bool())
}

func _num2Letter(num int) string {
	letterA := rune('A')
	num = int(letterA) + num
	return strings.ToUpper(string(rune(num)))
}

//jsoniter + xlsx
func _xlsx() bool {
	var file *xlsx.File
	var sheet2 *xlsx.Sheet
	var row2 *xlsx.Row
	var cell *xlsx.Cell
	var err error

	//var json = jsoniter.ConfigCompatibleWithStandardLibrary
	jsonStr := loadJson()
	sheets := jsoniter.Get([]byte(jsonStr), "data")

	//Any sheets
	//json := jsoniter.ConfigCompatibleWithStandardLibrary
	//json.Unmarshal([]byte(jsonStr), &sheets)
	sheetNames := sheets.Keys()
	if len(sheetNames) <= 0 {
		return false
	}

	file = xlsx.NewFile()
	startTime := time.Now().Unix()
	fmt.Println("start=", startTime)

	for _, sheetName := range sheetNames {
		sheet2, err = file.AddSheet(sheetName)
		if err != nil {
			fmt.Printf(err.Error())
			return false
		}

		rows := sheets.Get(sheetName, "rows")
		rowNum := rows.Size()
		fmt.Println(rowNum)
		for i := 0; i < rowNum; i++ {
			row := rows.Get(i) //jsoniter.Get([]byte(jsonStr), "data", val, i)
			colNum := row.Size()
			fmt.Println("row", i+1, time.Now().Unix())
			fmt.Println(colNum)
			row2 = sheet2.AddRow()
			for j := 0; j < colNum; j++ {
				col := row.Get(j)
				colNames := col.Keys()
				for _, colName := range colNames {
					switch colName {
					case "value":
						//fmt.Println(col.Get(colName).ToString())
						cell = row2.AddCell()
						cell.Value = col.Get(colName).ToString()
					case "format":
						//fmt.Println("format=", col.Get("value").ToString())
					}
				}
			}
		}
	}

	err = file.Save("MyXLSXFile.xlsx")
	fmt.Println("time elapsed:", (time.Now().Unix() - startTime))
	if err != nil {
		fmt.Printf(err.Error())
	} else {
		fmt.Print("done")
	}
	return true
}

//simpleJson + excelize
func _simple() bool {
	js, err := simplejson.NewJson([]byte(loadJson()))

	if err != nil {
		panic(err.Error())
	}

	data := js.Get("data").MustMap()

	xlsx := excelize.NewFile()
	var index int

	startTime := time.Now().Unix()
	if debug == 1 {
		fmt.Println("start=", time.Now().Unix())
	}

	var sheetNum = 0
	for sheetName, _ := range data {
		if sheetNum == 0 {
			index = xlsx.GetActiveSheetIndex()
			oldName := xlsx.GetSheetName(index)
			xlsx.SetSheetName(oldName, sheetName)
		} else {
			index = xlsx.NewSheet(sheetName)
		}
		sheetNum++

		formats, _ := js.Get("data").Get(sheetName).Get("col_format").Array()
		//os.Exit(1)
		for i := 0; i < len(formats); i++ {
			format := js.Get("data").Get(sheetName).Get("col_format").GetIndex(i)
			colStr := format.Get("col").MustString()
			colWidth := format.Get("width").MustFloat64()
			if debug == 1 {
				fmt.Println("format size2", colStr, colWidth)
			}
			//os.Exit(2)
			if colStr != "" && colWidth > 0 {
				cols := strings.Split(colStr, ":")
				xlsx.SetColWidth(sheetName, cols[0], cols[1], colWidth)
			} else if debug == 1 {
				fmt.Println("col format error")
			}
		}

		rows, _ := js.Get("data").Get(sheetName).Get("rows").Array()
		for i := 0; i < len(rows); i++ {
			if debug == 1 {
				fmt.Println("row", i+1, float64(time.Now().UnixNano())/1000000000)
			}
			excelRowNum := strconv.Itoa(i + 1)
			row, _ := js.Get("data").Get(sheetName).Get("rows").GetIndex(i).Array()
			for j := 0; j < len(row); j++ {
				col := js.Get("data").Get(sheetName).Get("rows").GetIndex(i).GetIndex(j)
				cellName := _num2Letter(j) + "" + excelRowNum
				xlsx.SetCellValue(sheetName, cellName, col.Get("value").MustString())

				colFormat := col.Get("format")
				if colFormat != nil {
					var styleString = ""
					val := colFormat.Get("bg_color").MustString()
					if val != "" {
						styleString = styleString + "\"fill\":{\"type\":\"pattern\",\"pattern\":1,\"color\":[\"" + val + "\"]},"
					}
					val = colFormat.Get("align").MustString()
					if val != "" {
						val2 := colFormat.Get("valign").MustString()
						if val2 != "" {
							styleString = styleString + strings.Join([]string{"\"alignment\":{\"horizontal\":\"", val + "\",\"vertical\":\"", val2, "\"}"}, "")
						} else {
							styleString = styleString + strings.Join([]string{"\"alignment\":{\"horizontal\":\"", val, "\"}"}, "")
						}
					} else {
						val = colFormat.Get("valign").MustString()
						if val != "" {
							styleString = styleString + strings.Join([]string{"\"alignment\":{\"vertical\":\"", val, "\"}"}, "")
						}
					}

					if styleString != "" {
						style, err := xlsx.NewStyle("{" + styleString + "}")
						if err == nil {
							xlsx.SetCellStyle(sheetName, cellName, cellName, style)
						} else {
							fmt.Println(err)
						}
					}
				}
			}
			if i > 1000 {
				break
			}
		}
	}

	saveFile := common.File_dir(jsonFile) + "/" + strconv.FormatInt(common.Rand_int(10000, 99999), 10) + ".xlsx"
	err = xlsx.SaveAs(saveFile)
	if err != nil {
		if debug == 1 {
			fmt.Println(saveFile, err)
		}
	} else {
		fmt.Println(saveFile)
	}

	if debug == 1 {
		fmt.Println("end=", time.Now().Unix()-startTime)
	}
	return true

}

//jsoniter + excelize
func _jsoniter() bool {
	//var json = jsoniter.ConfigCompatibleWithStandardLibrary
	jsonStr := loadJson()
	sheets := jsoniter.Get([]byte(jsonStr), "data")

	//Any sheets
	sheetNames := sheets.Keys()
	if len(sheetNames) <= 0 {
		return false
	}

	xlsx := excelize.NewFile()
	var index int

	startTime := time.Now().Unix()
	fmt.Println("start=", startTime)

	for key, sheetName := range sheetNames {
		if key == 0 {
			index = xlsx.GetActiveSheetIndex()
			oldName := xlsx.GetSheetName(index)
			xlsx.SetSheetName(oldName, sheetName)
		} else {
			index = xlsx.NewSheet(sheetName)
		}

		xlsx.SetActiveSheet(index)
		rows := sheets.Get(sheetName, "rows")
		rowNum := rows.Size()
		//fmt.Println(rowNum)
		for i := 0; i < rowNum; i++ {
			row := rows.Get(i)
			colNum := row.Size()
			fmt.Println("row", i+1, time.Now().Unix())
			//fmt.Println(colNum)
			for j := 0; j < colNum; j++ {
				col := row.Get(j)
				colNames := col.Keys()
				//fmt.Println(len(colNames))
				for _, colName := range colNames {
					excelRowNum := strconv.Itoa(i + 1)
					switch colName {
					case "value":
						//fmt.Println("col=", strings.Join([]string{_num2Letter(j), excelRowNum}, ""))
						xlsx.SetCellValue(sheetName, _num2Letter(j)+""+excelRowNum, col.Get(colName).ToString())
						//fmt.Println("value=", col.Get("value").ToString())
					case "format":
						//fmt.Println("format=", col.Get("value").ToString())
					}
				}
			}
			//fmt.Println(col.Get);
			if i > 1000 {
				break
			} //break
		}
	}

	err := xlsx.SaveAs(common.File_dir(jsonFile) + "./Book1.xlsx")
	if err != nil {
		fmt.Println(err, index)
	} else {
		fmt.Println("done")
	}
	if debug == 1 {
		fmt.Println("end=", time.Now().Unix()-startTime)
	}
	return true
}

func main() {
	flag.StringVar(&jsonFile, "jsonfile", "", "the full path of the json file")
	flag.IntVar(&debug, "debug", 1, "show debug info or not")
	flag.IntVar(&jsonMethod, "method", 1, "json lib to call")
	flag.Parse()

	if jsonFile == "" {
		fmt.Print("empty argv", jsonFile)
		os.Exit(0)
	} else if debug == 1 {
		fmt.Println("arg=", jsonFile)
	}

	if !common.File_exists(jsonFile) {
		fmt.Print("not found>>", jsonFile)
		os.Exit(0)
	}

	switch jsonMethod {
	case 1:
		_simple()
	case 2:
		_jsoniter()
	}

}
