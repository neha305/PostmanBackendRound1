package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	f, err := excelize.OpenFile("Sample-Menu.xlsx")
	i := 65
	clNo := 65
	styleID := f.GetCellStyle("Sheet1", "A1")
	//Changing style of the cells containing date to regular text format
	for {
		cellValue := f.GetCellValue("Sheet1", fmt.Sprintf("%s%d", string(i), 2))
		if cellValue == "" {
			clNo = i - 1
			break
		}
		f.SetCellStyle("Sheet1", fmt.Sprintf("%s%d", string(i), 2), fmt.Sprintf("%s%d", string(i), 2), styleID)
		i++
	}
	if err != nil {
		log.Fatal(err)
	}
	rows := f.GetRows("Sheet1")
	maxRow := len(rows)
	sheetName := "Sheet1"
	var day, meal, item string
	var Items []string

	//Calling Function 1
	fmt.Print("Enter day and meal in all caps: ")
	fmt.Scanln(&day)
	fmt.Scanln(&meal)
	Items = f1(f, day, meal, maxRow, sheetName)
	fmt.Println(strings.Join(Items, ","))

	//Calling Function 2
	fmt.Print("Enter day and meal in all caps: ")
	fmt.Scanln(&day)
	fmt.Scanln(&meal)
	num := f2(f, day, meal, maxRow, sheetName)
	fmt.Println("Number of items: ", num)

	//Calling Function 3
	fmt.Print("Enter day, meal and item in all caps: ")
	fmt.Scanln(&day)
	fmt.Scanln(&meal)
	reader := bufio.NewReader(os.Stdin)
	item, err1 := reader.ReadString('\n')
	if err1 != nil {
		log.Fatal(err1)
	}
	item = strings.Replace(item, "\n", "", -1)
	f3(f, day, meal, maxRow, sheetName, item)

	//Calling Function 4
	f4(f, maxRow, sheetName, clNo)
	fmt.Println("And saved.")

	//Calling Function 5
	f5()
	fmt.Println("\n\nFin.")
}

func f1(f *excelize.File, day string, meal string, maxRow int, sheetName string) []string {
	var Items []string
	result := f.SearchSheet("Sheet1", day)
	column := string(result[0][0])

	r := 1
	for i := 1; i <= maxRow; i++ {
		cellValue := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", column, i))
		if cellValue == meal {
			r = i
			break
		}
	}

	for i := r + 1; i <= maxRow; i++ {
		cellValue := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", column, i))
		if (cellValue == day) || (cellValue == "") {
			break
		}
		Items = append(Items, cellValue)
	}
	return Items
}

func f2(f *excelize.File, day string, meal string, maxRow int, sheetName string) int {
	var Items []string
	Items = f1(f, day, meal, maxRow, sheetName)
	num := len(Items)
	return num
}

func f3(f *excelize.File, day string, meal string, maxRow int, sheetName string, item string) {
	var Items []string
	Items = f1(f, day, meal, maxRow, sheetName)
	i := 0
	for _, food := range Items {
		if food == item {
			i = 1
			break
		}
	}
	if i == 1 {
		fmt.Println("Item is found.")
	} else {
		fmt.Println("Item is not found.")
	}
}

func f4(f *excelize.File, maxRow int, sheetName string, clNo int) {
	jsonData := make([]map[string]string, 0)
	var Items []string
	meals := []string{"BREAKFAST", "LUNCH", "DINNER"}
	for i := 65; i <= clNo; i++ {
		column := string(i)
		day := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", column, 1))
		date := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", column, 2))
		for _, meal := range meals {
			data := make(map[string]string)
			Items = f1(f, day, meal, maxRow, sheetName)
			data["day"] = day
			data["date"] = date
			data["meal"] = meal
			data["items"] = strings.Join(Items, ",")
			jsonData = append(jsonData, data)
		}
	}
	jsonBytes, err := json.MarshalIndent(jsonData, "", "    ")
	if err != nil {
		log.Fatal(err)
	}

	jsonFile, err := os.Create("output.json")
	if err != nil {
		log.Fatal(err)
	}
	defer jsonFile.Close()

	_, err = jsonFile.Write(jsonBytes)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Excel file converted to JSON.")
}

func f5() {
	type Food struct {
		Day   string `json:"day"`
		Date  string `json:"date"`
		Meal  string `json:"meal"`
		Items string `json:"items"`
	}
	file, err := os.Open("output.json")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}
	defer file.Close()

	var foods []Food
	if err := json.NewDecoder(file).Decode(&foods); err != nil {
		fmt.Println("Error decoding JSON:", err)
		return
	}

	fmt.Println("Structs:")
	for _, p := range foods {
		fmt.Printf("Day: %s, \nDate: %s, \nMeal: %s, \nItems: %s\n\n", p.Day, p.Date, p.Meal, p.Items)
	}
}
