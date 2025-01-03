package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	ical "github.com/arran4/golang-ical"
	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
)

var FilePath string
var GroupName string
var StartingColumn string
var StartingRow int

func main() {
	flag.StringVar(&FilePath, "path", "", "Spreadsheet file path")
	flag.StringVar(&GroupName, "group", "", "Class group")
	flag.StringVar(&StartingColumn, "sc", "B", "Starting column")
	flag.IntVar(&StartingRow, "sr", 1, "Starting row")

	flag.Parse()

	f, err := excelize.OpenFile(FilePath)

	if err != nil {
		fmt.Println(err)
		return
	}

	defer f.Close()

	sheet := f.GetSheetList()[0]

	sections, err := ParseSections(f, sheet)

	if err != nil {
		fmt.Println(err)
		return
	}

	dateColumn := ""

	{
		val, err := excelize.ColumnNameToNumber(StartingColumn)

		if err != nil {
			fmt.Println(err)
			return
		}

		dateColumn, err = excelize.ColumnNumberToName(val + 4)

		if err != nil {
			fmt.Println(err)
			return
		}
	}

	startingDate, err := f.GetCellValue(sheet, fmt.Sprintf("%s%d", dateColumn, StartingRow+2))

	if err != nil {
		fmt.Print(err)
		return
	}

	classStart, err := time.ParseInLocation("01-02-06", startingDate, time.Local)

	if err != nil {
		fmt.Print(err)
		return
	}

	chosenSection := sections[GroupName]

	currentCol := 2

	currentDay := classStart.Day()

	classMap := make(map[int][]Class, 0)

	size, err := f.GetSheetDimension(sheet)

	if err != nil {
		fmt.Println(err)
		return
	}

	colSize, _, err := excelize.CellNameToCoordinates(strings.Split(size, ":")[1])

	if err != nil {
		fmt.Println(err)
		return
	}

	for {
		for i := 1; i < 16; i++ {
			cell, err := excelize.CoordinatesToCellName(currentCol+i, chosenSection.SectionStart+2)

			if err != nil {
				fmt.Println(err)
				return
			}

			cellValue, err := f.GetCellValue(sheet, cell)

			if err != nil {
				fmt.Println(err)
				return
			}

			if cellValue == "" {
				continue
			}

			teacherCell, err := excelize.CoordinatesToCellName(currentCol+i, chosenSection.SectionStart+3)

			if err != nil {
				fmt.Println(err)
				return
			}

			teacherValue, err := f.GetCellValue(sheet, teacherCell)

			if err != nil {
				fmt.Println(err)
				return
			}

			weekCell, err := excelize.CoordinatesToCellName(currentCol+i, chosenSection.SectionStart+4)

			if err != nil {
				fmt.Println(err)
				return
			}

			weekValue, err := f.GetCellValue(sheet, weekCell)

			if err != nil {
				fmt.Println(err)
				return
			}

			weeks := strings.Split(weekValue, ",")

			classWeeks := make([]int, 0)

			for _, v := range weeks {
				if strings.Contains(v, "/") || len(v) == 0 {
					continue
				}

				week, err := strconv.Atoi(v)

				if err != nil {
					fmt.Println(err)
					return
				}

				classWeeks = append(classWeeks, week)
			}

			locations := make([]string, 0)

			for idx := 3; idx < 8; idx++ {
				locationCell, err := excelize.CoordinatesToCellName(currentCol+i+1, chosenSection.SectionStart+idx)

				if err != nil {
					fmt.Println(err)
					return
				}

				locationValue, err := f.GetCellValue(sheet, locationCell)

				if err != nil {
					fmt.Println(err)
					return
				}

				if locationValue == "" {
					continue
				} else {
					locations = append(locations, locationValue)
				}

			}

			class := Class{
				Name:      cellValue,
				Teacher:   teacherValue,
				Weeks:     classWeeks,
				Start:     i,
				End:       i + 1,
				Locations: locations,
			}

			if _, ok := classMap[currentDay]; !ok {
				classMap[currentDay] = []Class{class}
			} else {
				classMap[currentDay] = append(classMap[currentDay], class)
			}
		}

		currentCol += 17
		currentDay++

		if currentCol >= colSize {
			break
		}
	}

	cal := ical.NewCalendar()
	cal.SetTimezoneId("Europe/Warsaw")
	cal.SetXWRTimezone("Europe/Warsaw")

	for classDay, classList := range classMap {
		for _, class := range classList {
			evt := cal.AddEvent(strings.ReplaceAll(uuid.NewString(), "-", ""))

			evt.SetSummary(class.Name)
			evt.SetDescription(fmt.Sprintf("Prowadzący: %s", class.Teacher))

			evt.SetLocation(strings.TrimSpace(strings.Join(class.Locations, ", ")))

			evt.SetCreatedTime(time.Now())
			evt.SetDtStampTime(time.Now())
			evt.SetModifiedAt(time.Now())

			evt.SetStartAt(GetClass(classStart, classDay, class.Start, true))
			evt.SetEndAt(GetClass(classStart, classDay, class.End, false))

			evt.SetProperty("VTIMEZONE", "Europe/Warsaw")
		}
	}

	os.WriteFile("calendar.ics", []byte(cal.Serialize()), 0644)
}

type Class struct {
	Name      string
	Teacher   string
	Weeks     []int
	Locations []string
	Start     int
	End       int
}

type Section struct {
	SectionName  string
	SectionStart int
	SectionEnd   int
}

func ParseSections(f *excelize.File, sheet string) (map[string]Section, error) {
	currentRow := StartingRow
	startRow := currentRow

	sections := make(map[string]Section)

	tempSection := make([]string, 0)

	for {
		idx, err := f.GetCellStyle(sheet, fmt.Sprintf("%s%d", StartingColumn, currentRow))

		if err != nil {
			return nil, err
		}

		cellValue, err := f.GetCellValue(sheet, fmt.Sprintf("%s%d", StartingColumn, currentRow))

		if err != nil {
			return nil, err
		}

		style, err := f.GetStyle(idx)

		if err != nil {
			return nil, err
		}

		tempSection = append(tempSection, strings.TrimSpace(cellValue))

		for _, v := range style.Border {
			if v.Type == "top" || v.Type == "bottom" {
				tempStr := strings.TrimSpace(strings.Join(tempSection, " "))

				if tempStr != "" {

					tempVal, err := strconv.Atoi(tempStr)

					if err != nil && strconv.Itoa(tempVal) != tempStr {
						sections[tempStr] = Section{
							SectionName:  tempStr,
							SectionStart: startRow - 2,
							SectionEnd:   currentRow,
						}

						startRow = currentRow
					}
				} else {
					startRow = currentRow
				}

				tempSection = make([]string, 0)
			}
		}

		currentRow++

		if currentRow >= 140 {
			break
		}
	}

	return sections, nil
}

func GetClass(base time.Time, day, hour int, isStart bool) time.Time {
	var classTime time.Time

	if isStart {
		classTime = HourToTimeStart[hour]
	} else {
		classTime = HourToTimeEnd[hour]
	}

	// Since some importer apps don't support timezone, we need to subtract 2 hours from the time
	return time.Date(base.Year(), base.Month(), day, classTime.Hour()-2, classTime.Minute(), classTime.Second(), classTime.Nanosecond(), time.UTC)
}

var HourToTimeStart = map[int]time.Time{
	1:  time.Date(0, 0, 0, 8, 0, 0, 0, time.Local),
	2:  time.Date(0, 0, 0, 8, 45, 0, 0, time.Local),
	3:  time.Date(0, 0, 0, 9, 40, 0, 0, time.Local),
	4:  time.Date(0, 0, 0, 10, 25, 0, 0, time.Local),
	5:  time.Date(0, 0, 0, 11, 30, 0, 0, time.Local),
	6:  time.Date(0, 0, 0, 12, 15, 0, 0, time.Local),
	7:  time.Date(0, 0, 0, 13, 10, 0, 0, time.Local),
	8:  time.Date(0, 0, 0, 13, 55, 0, 0, time.Local),
	9:  time.Date(0, 0, 0, 14, 45, 0, 0, time.Local),
	10: time.Date(0, 0, 0, 15, 30, 0, 0, time.Local),
	11: time.Date(0, 0, 0, 16, 20, 0, 0, time.Local),
	12: time.Date(0, 0, 0, 17, 5, 0, 0, time.Local),
	13: time.Date(0, 0, 0, 17, 55, 0, 0, time.Local),
	14: time.Date(0, 0, 0, 18, 40, 0, 0, time.Local),
	15: time.Date(0, 0, 0, 19, 30, 0, 0, time.Local),
	16: time.Date(0, 0, 0, 20, 15, 0, 0, time.Local),
}

var HourToTimeEnd = map[int]time.Time{
	1:  time.Date(0, 0, 0, 8, 45, 0, 0, time.Local),
	2:  time.Date(0, 0, 0, 9, 30, 0, 0, time.Local),
	3:  time.Date(0, 0, 0, 10, 25, 0, 0, time.Local),
	4:  time.Date(0, 0, 0, 11, 10, 0, 0, time.Local),
	5:  time.Date(0, 0, 0, 12, 15, 0, 0, time.Local),
	6:  time.Date(0, 0, 0, 13, 00, 0, 0, time.Local),
	7:  time.Date(0, 0, 0, 13, 55, 0, 0, time.Local),
	8:  time.Date(0, 0, 0, 14, 40, 0, 0, time.Local),
	9:  time.Date(0, 0, 0, 15, 30, 0, 0, time.Local),
	10: time.Date(0, 0, 0, 16, 15, 0, 0, time.Local),
	11: time.Date(0, 0, 0, 17, 5, 0, 0, time.Local),
	12: time.Date(0, 0, 0, 17, 50, 0, 0, time.Local),
	13: time.Date(0, 0, 0, 18, 40, 0, 0, time.Local),
	14: time.Date(0, 0, 0, 19, 25, 0, 0, time.Local),
	15: time.Date(0, 0, 0, 20, 15, 0, 0, time.Local),
	16: time.Date(0, 0, 0, 21, 0, 0, 0, time.Local),
}
