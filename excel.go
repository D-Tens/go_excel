package go_excel

import (
	"errors"
	"fmt"
	"github.com/tealeg/xlsx"
	"reflect"
	"strconv"
	"strings"
)

type Excel struct {
	File *xlsx.File
}

type Sheet struct {
	Sheet            *xlsx.Sheet
	AutoCreateHeader bool
}

func NewExcel() *Excel {
	return &Excel{
		File: xlsx.NewFile(),
	}
}

func GetElement(e reflect.Value) reflect.Value {
	for e.Kind() == reflect.Ptr || e.Kind() == reflect.Interface {
		e = e.Elem()
	}
	return e
}

func GetAssertString(sData interface{}) string {
	if s, ok := sData.(string); ok {
		return s
	} else if vInt, ok := sData.(int); ok {
		return strconv.Itoa(vInt)
	} else if vInt, ok := sData.(int32); ok {
		return strconv.FormatInt(int64(vInt), 10)
	} else if vInt, ok := sData.(float64); ok {
		return strconv.FormatFloat(vInt, 'f', -1, 64)
	} else if vInt, ok := sData.(int64); ok {
		return strconv.FormatInt(vInt, 10)
	}
	return ""
}

func (s *Sheet) SetAutoCrateHeader(isTrue bool) {
	s.AutoCreateHeader = isTrue
}

func (e *Excel) SaveExcel(fPath string) error {
	fPath = fmt.Sprintf("%s.xlsx", strings.TrimRight(fPath, ".xlsx"))
	if err := e.File.Save(fPath); err != nil {
		return err
	}
	return nil
}

func (e *Excel) AddSheet(name string) (*Sheet, error) {
	if sheet, err := e.File.AddSheet(name); err != nil {
		return nil, err
	} else {
		return &Sheet{
			Sheet:            sheet,
			AutoCreateHeader: true,
		}, nil
	}
}

func (s *Sheet) SetTitle(title string, sType interface{}) *Sheet {
	row := s.Sheet.AddRow()
	cell := row.AddCell()
	cell.Merge(reflect.TypeOf(sType).NumField()-1, 0)
	style := xlsx.NewStyle()
	style.Alignment = xlsx.Alignment{Horizontal: "center"}
	cell.SetStyle(style)
	cell.SetString(title)
	return s
}

func (s *Sheet) SetHeader(sData interface{}) error {
	fmt.Println(sData)
	vData := reflect.ValueOf(sData)
	if vData.Kind() != reflect.Slice || vData.Len() == 0 {
		return nil
	}
	hValue := GetElement(vData.Index(0))
	header := make([]string, 0)
	fmt.Println("---", hValue.Kind(), reflect.Slice,hValue)
	switch hValue.Kind() {
	case reflect.Struct:
		hType := hValue.Type()
		for i := 0; i < hType.NumField(); i++ {
			header = append(header, hType.Field(i).Tag.Get("xlsx"))
		}
	case reflect.Slice:
		for i := 0; i < hValue.Len(); i++ {
			header = append(header, GetAssertString(hValue.Index(i).Interface()))
		}
	default:
		return errors.New(fmt.Sprintf("Must Is Struct OR Slice..."))
	}
	row := s.Sheet.AddRow()
	for _, k := range header {
		cell := row.AddCell()
		cell.SetString(k)
	}
	return nil
}

func (s *Sheet) AddData(sData interface{}) error {
	dType := reflect.TypeOf(sData)
	if dType.Kind() != reflect.Slice {
		return errors.New(fmt.Sprintf("Must is Slice..."))
	}
	if s.AutoCreateHeader {
		if err := s.SetHeader(sData); err != nil {
			return errors.New(fmt.Sprintf("Create Tables Header Error..."))
		}
	}
	sheet := s.Sheet
	vData := reflect.ValueOf(sData)
	if vData.Len() == 0 {
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Merge(2, 0)
		cell.SetString("Not Data...")
	}
	for k := 0; k < vData.Len(); k++ {
		v := GetElement(vData.Index(k))
		if s.AutoCreateHeader && v.Kind() == reflect.Slice && k == 0 {
			continue
		}
		row := sheet.AddRow()
		switch v.Kind() {
		case reflect.Struct:
			for i := 0; i < v.NumField(); i++ {
				cell := row.AddCell()
				field := v.Field(i)
				switch field.Kind() {
				case reflect.String:
					cell.SetString(field.String())
					cell.NumFmt = "@"
				case reflect.Int, reflect.Int32, reflect.Int64:
					cell.SetInt64(int64(field.Int()))
				case reflect.Float32, reflect.Float64:
					cell.SetFloat(float64(field.Float()))
				default:
					cell.SetValue(GetAssertString(field.Interface()))
				}
			}
		case reflect.Slice:
			for i := 0; i < v.Len(); i++ {
				cell := row.AddCell()
				field := v.Index(i)
				switch field.Kind() {
				case reflect.String:
					cell.SetString(field.String())
					cell.NumFmt = "@"
				case reflect.Int, reflect.Int32, reflect.Int64:
					cell.SetInt64(int64(field.Int()))
				case reflect.Float32, reflect.Float64:
					cell.SetFloat(float64(field.Float()))
				default:
					cell.SetValue(GetAssertString(field.Interface()))
				}
			}
		default:
			return errors.New(fmt.Sprintf("Kind Must is Struct Or Slice..."))
		}
	}
	return nil
}
