package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"

	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
	"gopkg.in/ini.v1"
)

var (
	version = "20221220"
	//nolint:gochecknoglobals // TODO: уйти от глобальной переменной
	DEBUG = false
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
	style   _style
	styleID int
}

type _column struct {
	id int
	// name        string
	width      int
	replaces   []_replace
	finds      []_find
	deleted    bool
	horizontal string
	size       float64
}

type _defColParam struct {
	width int
	// horizontal string
}

type _style struct {
	Alignment excelize.Alignment
	Border    []excelize.Border
	Font      excelize.Font
	Fill      excelize.Fill
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
	horizontal string
}
type _cfg struct {
	sheetName string
	title     _title
	header    _header
	delimiter rune
	cols      []_column
	style     _style
	defcol    _defColParam
}

func usage() {
	fmt.Println("Конвертирование csv в xlsx с форматированием.")
	fmt.Printf("версия: %s\n\n", version)
	fmt.Println("csv2xlsx [command] <args>")
	fmt.Println("command:")
	flag.PrintDefaults()

}

// Парсим проблемые поля из csv
func csvParseErrFieldCount(reads []string, fieldCount int) [][]string {
	// Массив с распарсенными по полям строками
	var parsedRows [][]string
	// id поля которое парсится
	var unCol = 0
	// Текущая строка csv
	var st []string

	// Идем по полученным полям и ищем косяк c "":
	// "" будет как "\n", при этом два поля объеденятся
	for _, field := range reads {
		f := strings.Split(field, "\n")

		for i, r := range f {
			// все что больше 0, исправленные. В начале надо убрать "
			if i != 0 {
				r = strings.Replace(r, "\"", "", 1)
			}
			fmt.Println("[DBG] ", r)

			if len(r) != 0 {
				st = append(st, r)
				unCol++
			}
			if unCol >= fieldCount {
				parsedRows = append(parsedRows, st)
				fmt.Printf("[DBG] add rows, %d columns\n", len(st))
				st = make([]string, 0)
				unCol = 0
			}
		}
	}
	return parsedRows
}

// Генерация xlsx из csv
func generateXLSXFromCSV(csvPath string, delimiter rune) (*excelize.File, error) {
	if DEBUG {
		fmt.Println("[DBG] Открываем файл ", csvPath)
	}
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return nil, err
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	// reader.Comma = '\t'
	reader.Comma = delimiter
	reader.LazyQuotes = true
	// reader.ReuseRecord = true

	xlsxFile := excelize.NewFile()

	// sheet := "Sheet1"
	sheet := xlsxFile.GetSheetName(0)

	row := 1

	for {
		if DEBUG {
			fmt.Printf("[DBG] Обрабатываем строку: %d\n", row)
		}
		// Считываем следующую строку
		fields, err := reader.Read()
		if errors.Is(err, io.EOF) {
			break
		}
		//nolint:nestif // TODO: переписать
		if err != nil {
			switch {
			// Пытаемся обработать заголовок ВНИИРА
			case errors.Is(err, csv.ErrFieldCount) && row == 2:
				if DEBUG {
					fmt.Println("[DBG] не совпало кол-во полей с пред. записью.\nТ.к. срока 2, считаем заголовком.")
				}
				reader.FieldsPerRecord = 0

			case errors.Is(err, csv.ErrFieldCount):

				// TODO: переписать по человечески
				if DEBUG {
					fmt.Println("[ERR] не совпало кол-во полей с пред. записью.")
					fmt.Println(fields)
					fmt.Println("[DBG] пытаемся обработать")
				}

				// Пытаемся обработать ""
				parsedRows := csvParseErrFieldCount(fields, reader.FieldsPerRecord-1)

				// Добавляем полученные строки в файл
				for _, fields := range parsedRows {
					for i, field := range fields {
						if DEBUG {
							fmt.Printf("[DBG] строка: %d, столбец: %d\n", row, i)
						}
						column, err := excelize.ColumnNumberToName(i + 1)
						if err != nil {
							return nil, err
						}

						err = xlsxFile.SetCellStr(sheet, fmt.Sprintf("%s%d", column, row), field)
						if err != nil {
							return nil, err
						}
					}

					row++
				}
				continue
			default:
				return nil, err
			}
		}

		for i, field := range fields {
			if DEBUG {
				fmt.Printf("[DBG] строка: %d, столбец: %d\n", row, i)
			}
			column, err := excelize.ColumnNumberToName(i + 1)
			if err != nil {
				return nil, err
			}

			err = xlsxFile.SetCellStr(sheet, fmt.Sprintf("%s%d", column, row), field)
			if err != nil {
				return nil, err
			}
		}

		row++
	}

	return xlsxFile, nil
}

func main() {
	var xlsxPath = flag.String("o", "", "Path to the XLSX output file")
	var csvPath = flag.String("f", "", "Path to the CSV input file")
	var cfgPatch = flag.String("c", "", "Path to the config file (optional)")
	var delimiter = flag.String("d", "\t", "Delimiter for felds in the CSV input (optional)")
	flag.BoolVar(&DEBUG, "debug", false, "Debug (optional)")

	if len(os.Args) < 2 {
		usage()
		return
	}

	flag.Usage = usage

	flag.Parse()
	if *xlsxPath == "" || *csvPath == "" {
		flag.Usage()
		return
	}

	cfg := new(_cfg)

	// Загрузка настроек
	if *cfgPatch != "" {
		if DEBUG {
			fmt.Println("Load config: ", *cfgPatch)
		}
		if err := loadCFG(*cfgPatch, cfg); err != nil {
			fmt.Println(err.Error())
		}
	} else {
		cfg.delimiter = delimiterToRune(*delimiter)
	}

	// Генерируем xlsx
	xlsxFile, err := generateXLSXFromCSV(*csvPath, cfg.delimiter)
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	// Если есть конфиг, применяем настройки
	if *cfgPatch != "" {
		if err := applyFormatting(xlsxFile, cfg); err != nil {
			fmt.Println(err.Error())
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
func loadCFG(iniFile string, cfg *_cfg) error {
	inifile, err := ini.ShadowLoad(iniFile)
	if err != nil {
		return err
	}

	// Общие настройки
	cfg.sheetName = inifile.Section("").Key("sheet").MustString("")
	cfg.delimiter = delimiterToRune(inifile.Section("").Key("delimiter").MustString(`\t`))

	// Добавить границу
	if inifile.Section("").Key("border").MustBool(false) {
		cfg.style.Border = []excelize.Border{
			{
				Type:  "left",
				Color: "#000000",
				Style: 2,
			}, {
				Type:  "top",
				Color: "#000000",
				Style: 2,
			}, {
				Type:  "bottom",
				Color: "#000000",
				Style: 2,
			}, {
				Type:  "right",
				Color: "#000000",
				Style: 2,
			},
		}
	}

	// Задать общий стиль

	// Перенос текста
	cfg.style.Alignment.WrapText = true

	// Alignment
	if err := setAlignment(&cfg.style.Alignment, "horizontal", inifile.Section("").Key("horizontal").MustString("")); err != nil {
		fmt.Println("[WRN]\tloadCFG: horizontal: ", err.Error())
	}

	if err := setAlignment(&cfg.style.Alignment, "vertical", inifile.Section("").Key("vertical").MustString("")); err != nil {
		fmt.Println("[WRN]\tloadCFG: vertical: ", err.Error())
	}

	// Font
	if inifile.Section("").Key("size").MustFloat64(0) != 0 {
		cfg.style.Font.Size = inifile.Section("").Key("size").MustFloat64(0)
	}
	if inifile.Section("").Key("family").MustString("") != "" {
		cfg.style.Font.Family = inifile.Section("").Key("family").MustString("")
	}

	cfg.defcol.width = inifile.Section("").Key("width").MustInt(-1)

	// Секции
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
				cfg.header.horizontal = inifile.Section(section).Key("horizontal").MustString("center")
			} else {
				cfg.header.enable = false
			}
		}

		// Настройки столбцов
		// Добавляем данные секции (столбца)
		col, err := loadColumnSettings(inifile, section)
		if err == nil {
			cfg.cols = append(cfg.cols, col)
		}

	} // конец цикла по секциям

	return nil
}

// Загрузка настроек столбца
func loadColumnSettings(inifile *ini.File, section string) (_column, error) {
	var col _column
	var err error

	col.id, err = strconv.Atoi(section)
	// Если секция не цифровая, пропускаем
	if err != nil {
		return col, err
	}

	col.width = inifile.Section(section).Key("width").MustInt(-1)
	col.horizontal = inifile.Section(section).Key("horizontal").MustString("")
	col.deleted = inifile.Section(section).Key("delete").MustBool(false)
	col.size = inifile.Section(section).Key("size").MustFloat64(0)

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
	return col, nil
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
		}
	}

	return rune('\t')
}

// Применяем настройки форматирования
func applyFormatting(xlsxFile *excelize.File, cfg *_cfg) error {

	// Получаем название листа
	sheetName := xlsxFile.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("sheet not found")
	}

	// задаем форматирование всей таблице
	if err := xlsxSetTableStyle(xlsxFile, sheetName, cfg.style, cfg.defcol); err != nil {
		return err
	}

	// обрабатываем параметры столбцов
	if err := xlsxSetColumnFormat(xlsxFile, sheetName, cfg.cols, cfg.style); err != nil {
		return err
	}

	// Стиль заголовка
	if cfg.header.enable {
		// Получаем последний столбец
		cols, err := xlsxFile.GetCols(sheetName)
		if err != nil {
			return err
		}
		lastColumn, err := excelize.ColumnNumberToName(len(cols))
		if err != nil {
			return err
		}
		// Устанавливаем стиль заголовка
		if err := xlsxSetHeader(xlsxFile, sheetName, fmt.Sprintf("A%d", cfg.header.row), fmt.Sprintf("%s%d", lastColumn, cfg.header.row), cfg.header, cfg.style); err != nil {
			return err
		}
	}

	// Добавить Title
	if cfg.title.enable {
		if err := xlsxAddTitle(xlsxFile, sheetName, cfg.title, cfg.style); err != nil {
			return err
		}
	}

	// переименовываем лист
	if cfg.sheetName != "" {
		if err := xlsxFile.SetSheetName(sheetName, cfg.sheetName); err != nil {
			return err
		}
	}

	return nil
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
