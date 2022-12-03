package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"

	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
	"gopkg.in/ini.v1"
)

type _replace struct {
	from string
	to   string
}

type _action struct {
	name  string
	value string
}

type _find struct {
	text    string
	actions []_action
	target  string
}

type _column struct {
	id int
	//name        string
	width    int
	replaces []_replace
	finds    []_find
	deleted  bool
}

type _style struct {
	alignment excelize.Alignment
	border    []excelize.Border
}

type _title struct {
	enable     bool
	text       string
	size       float64
	bold       bool
	color      string
	background string
}
type _header struct {
	enable     bool
	row        int
	size       float64
	bold       bool
	color      string
	background string
}
type _cfg struct {
	sheetName string
	title     _title
	border    bool
	header    _header
	delimiter rune
	cols      []_column
	style     _style
}

var cfg _cfg

func usage() {
	flag.PrintDefaults()
}

// Генерация xlsx из csv
func generateXLSXFromCSV(csvPath string, delimiter rune) (*excelize.File, error) {
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return nil, err
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	//reader.Comma = '\t'
	reader.Comma = delimiter

	xlsxFile := excelize.NewFile()

	//sheet := "Sheet1"
	sheet := xlsxFile.GetSheetName(0)

	row := 1

	for err == nil {
		// Считываем следующую строку
		fields, err := reader.Read()
		if err != nil {
			// switch err {
			// case csv.ErrFieldCount:
			// 	fmt.Println("Requested item not found")
			// case io.EOF:
			// 	fmt.Println("EOF")
			// default:
			// 	fmt.Println("Unknown error occurred")
			// }
			if err == io.EOF {
				break
			}
			// Пытаемся обработать заголовок ВНИИРА
			if row == 2 {
				reader.FieldsPerRecord = 0
			} else {
				return nil, err
			}

		}

		//if len(fields) != 0 {
		for i, field := range fields {
			column, err := excelize.ColumnNumberToName(i + 1)
			if err != nil {
				return nil, err
			}

			err = xlsxFile.SetCellStr(sheet, fmt.Sprintf("%s%d", column, row), field)
			if err != nil {
				return nil, err
			}
		}
		//}

		row++
	}

	return xlsxFile, nil
}

func main() {
	var xlsxPath = flag.String("o", "", "Path to the XLSX output file")
	var csvPath = flag.String("f", "", "Path to the CSV input file")
	var cfgPatch = flag.String("c", "", "Path to the config file (optional)")
	var delimiter = flag.String("d", "\t", "Delimiter for felds in the CSV input (optional)")

	if len(os.Args) < 2 {
		usage()
		return
	}

	flag.Parse()
	if *xlsxPath == "" || *csvPath == "" {
		usage()
		return
	}

	// Загрузка настроек
	if *cfgPatch != "" {
		if err := loadCFG(*cfgPatch); err != nil {
			fmt.Println(err.Error())
		}
	} else {
		cfg.delimiter = delimiterToRune(*delimiter)
	}

	xlsxFile, err := generateXLSXFromCSV(*csvPath, cfg.delimiter)
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	// Если есть конфиг, применяем насройки
	if *cfgPatch != "" {
		sheetName := xlsxFile.GetSheetName(0)
		if sheetName != "" {
			// Задать общий стиль
			xlsxSetDefaultStyle()

			// задаем форматирование
			if err := xlsxFormatSheet(xlsxFile, sheetName); err != nil {
				fmt.Println(err.Error())
			}

			// Добавить Title
			if cfg.title.enable {
				if err := xlsxAddTitle(xlsxFile, sheetName, cfg.title.text); err != nil {
					fmt.Println(err.Error())
				}
			}

			// переименовываем лист
			if cfg.sheetName != "" {
				xlsxFile.SetSheetName(sheetName, cfg.sheetName)
				//sheetName = cfg.sheetName
			}
		} else {
			fmt.Println(fmt.Errorf("sheet not found").Error())
		}
	}

	// Сохраняем результат
	if err := xlsxFile.SaveAs(*xlsxPath); err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	os.Exit(0)
}

// Загрузка настроек из ini в переменную cfg
func loadCFG(iniFile string) error {
	inifile, err := ini.ShadowLoad(iniFile)
	if err != nil {
		return err
	}

	// Общие настройки
	cfg.sheetName = inifile.Section("").Key("sheet").MustString("")
	cfg.border = inifile.Section("").Key("border").MustBool(false)
	cfg.delimiter = delimiterToRune(inifile.Section("").Key("delimiter").MustString(`\t`))

	for _, section := range inifile.SectionStrings() {
		// Кроме default
		if section == "DEFAULT" {
			continue
		}

		// Настройки шапки
		if section == "title" {
			if inifile.Section(section).Key("enable").MustBool(false) {
				cfg.title.enable = true
				cfg.title.text = inifile.Section(section).Key("text").MustString("")
				cfg.title.bold = inifile.Section(section).Key("bold").MustBool(false)
				cfg.title.size = inifile.Section(section).Key("size").MustFloat64(0)
				cfg.title.color = inifile.Section(section).Key("color").MustString("")
				cfg.title.background = inifile.Section(section).Key("background").MustString("")
			} else {
				cfg.title.enable = false
			}
		}

		// Настройки заголовка
		if section == "header" {
			if inifile.Section(section).Key("enable").MustBool(false) {
				cfg.header.enable = true
				cfg.header.row = inifile.Section(section).Key("row").MustInt(1)
				cfg.header.bold = inifile.Section(section).Key("bold").MustBool(false)
				cfg.header.size = inifile.Section(section).Key("size").MustFloat64(0)
				cfg.header.color = inifile.Section(section).Key("color").MustString("")
				cfg.header.background = inifile.Section(section).Key("background").MustString("")
			} else {
				cfg.header.enable = false
			}
		}

		// Настройки столбцов
		// Добавляем данные секции (столбца)
		var col _column
		col.id, err = strconv.Atoi(section)
		// Если секция не цифровая, пропускаем
		if err != nil {
			continue
		}
		col.width = inifile.Section(section).Key("width").MustInt(-1)
		col.deleted = inifile.Section(section).Key("delete").MustBool(false)

		// Load replaces
		allreplace := inifile.Section(section).Key("replace").ValueWithShadows()
		for _, repl := range allreplace {
			ss := strings.Split(repl, ",")
			if len(ss) == 2 {
				col.replaces = append(col.replaces, _replace{strings.Trim(ss[0], "\"'"), strings.Trim(ss[1], " \"'")})

			}
		}

		// Load finds
		allfinds := inifile.Section(section).Key("find").ValueWithShadows()
		for _, find := range allfinds {
			ss := strings.Split(find, ",")
			if len(ss) > 2 {
				var f _find
				f.text = strings.Trim(ss[0], " \"'")
				f.target = strings.Trim(ss[1], " \"'")
				// Идем по всем actions
				for i := 2; i < len(ss); i++ {
					actionsl := strings.Split(strings.TrimSpace(ss[i]), "=")
					var action _action
					action.name = strings.Trim(actionsl[0], "\"'")

					if len(actionsl) > 1 {
						action.value = strings.Trim(actionsl[1], "\"'")
					}
					f.actions = append(f.actions, action)

				}
				col.finds = append(col.finds, f)

			}
		}
		cfg.cols = append(cfg.cols, col)

	}

	return nil
}

func delimiterToRune(delimiter string) rune {
	switch delimiter {
	case ` `:
		return rune(' ')
	case `:`:
		return rune(':')
	case `;`:
		return rune(';')
	case `,`:
		return rune(',')
	case `\t`:
		return '\t'

	default:
		if len(delimiter) > 0 {
			return rune(delimiter[0])
		} else {
			return rune('\t')
		}
	}

	//return rune('\t')
}

// TO DO
// 1. Вынести xlsx в отделный модуль
// 2. Добавить условный формат:
// -eq
//     равно
// -ne
//     не равно
// -gt
//     больше
// -ge
//     больше или равно
// -lt
//     меньше
// -le
//     меньше или равно
