package main

import (
	"encoding/csv"
	"flag"
	"fmt"
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

type _cfg struct {
	sheetName string
	title     string
	border    bool
	header    bool
	delimiter rune
	cols      []_column
}

var cfg _cfg

func usage() {
	flag.PrintDefaults()
}

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
	fields, err := reader.Read()

	for err == nil {
		//if len(fields) != 0 {
		for i, field := range fields {
			column, err := excelize.ColumnNumberToName(i + 1)
			if err != nil {
				return nil, err
			}

			err = xlsxFile.SetCellStr(sheet, fmt.Sprintf("%s%d", column, row), field)
			if err != nil {
				fmt.Println(err.Error())
				return nil, err
			}
		}

		//}
		fields, err = reader.Read()
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

	sheetName := xlsxFile.GetSheetName(0)
	if sheetName != "" {
		// задаем форматирование
		if err := xlsxFormatSheet(xlsxFile, sheetName); err != nil {
			fmt.Println(err.Error())
		}

		// Добавить Title
		if cfg.title != "" {
			if err := xlsxAddTitle(xlsxFile, sheetName, cfg.title); err != nil {
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

	cfg.sheetName = inifile.Section("").Key("sheet").MustString("")
	cfg.title = inifile.Section("").Key("title").MustString("")
	cfg.border = inifile.Section("").Key("border").MustBool(false)
	cfg.header = inifile.Section("").Key("header").MustBool(false)

	cfg.delimiter = delimiterToRune(inifile.Section("").Key("delimiter").MustString(`\t`))

	for _, section := range inifile.SectionStrings() {
		// Кроме default
		if section != "DEFAULT" {
			// Добавляем данные секции (столбца)
			var col _column
			col.id, err = strconv.Atoi(section)
			if err != nil {
				return err
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
