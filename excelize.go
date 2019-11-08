package main

import (
	"flag"
	"fmt"
	"strconv"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func checkblank(str *string) bool {
	return strings.EqualFold(*str, "")
}

func CaseInsensitiveContains(s, substr string) bool {
	s, substr = strings.ToUpper(s), strings.ToUpper(substr)
	return strings.Contains(s, substr)
}

func main() {
	var wg sync.WaitGroup
	StartPtr := flag.String("start", "", "provide the starting column alphabet")
	EndPtr := flag.String("end", "", "provide the ending column alphabet")
	SheetPtr := flag.String("sheet", "", "provide the sheet")
	FilePtr := flag.String("file", "", "provide file name")
	SumPtr := flag.String("sumrow", "", "provide the sum row")
	flag.Parse()
	if checkblank(FilePtr) || checkblank(SheetPtr) || checkblank(StartPtr) || checkblank(EndPtr) || checkblank(SumPtr) {
		fmt.Println("Invalid Flags")
		return
	}
	//f, err := excelize.OpenFile("./Makerspace Spreadsheet.xlsx")
	f, err := excelize.OpenFile(*FilePtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	Start, err := excelize.ColumnNameToNumber(strings.ToUpper(*StartPtr))
	if err != nil {
		fmt.Println(err)
		return
	}
	End, err := excelize.ColumnNameToNumber(strings.ToUpper(*EndPtr))
	if err != nil {
		fmt.Println(err)
		return
	}
	//rows, err := f.GetRows("gradebook-export (2)")
	rows, err := f.GetRows(*SheetPtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	header := rows[0]
	rows = rows[1:]
	sumer := make(map[string][]int)
	location := make(map[string]int)
	for ind, val := range header[Start-1 : End] {
		temp := strings.Split(val, "-")
		tmp := strings.Trim(temp[1], " ")
		sumer[tmp] = make([]int, 0)
		location[tmp] = ind + Start
	}
	for ind, val := range header[End:] {
		for key := range sumer {
			if CaseInsensitiveContains(val, key) {
				sumer[key] = append(sumer[key], ind+End)
			}
		}
	}
	for ind, val := range rows {
		if len(header) == len(val) {
			wg.Add(1)
			go func(file *excelize.File, i int, sinrow []string) {
				defer wg.Done()
				catcher := make(map[string]int)
				for key, val := range sumer {
					for _, elem := range val {
						if num, err := strconv.Atoi(sinrow[elem]); err == nil {
							catcher[key] += num
						}
					}
					col, err := excelize.ColumnNumberToName(location[key])
					if err != nil {
						fmt.Println(err)
					}
					temp := strconv.Itoa(i)
					err = file.SetCellValue(*SheetPtr, col+temp, catcher[key])
					if err != nil {
						fmt.Println(err)
					}
					err = file.SetCellFormula(*SheetPtr, *SumPtr+temp, "=SUM("+*StartPtr+temp+":"+*EndPtr+temp+")")
					if err != nil {
						fmt.Println(err)
					}
				}
			}(f, ind+2, val)
		} else {
			continue
		}
	}
	wg.Wait()
	err = f.Save()
	if err != nil {
		fmt.Println("Unable to Process:", err)
	}
}
