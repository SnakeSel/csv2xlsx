package main

import (
	"fmt"

	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Задать общий стиль
func xlsxSetDefaultStyle() {
	// выравнивание текста
	cfg.style.alignment.WrapText = true
	cfg.style.alignment.Vertical = "center"

	// Добавить границу
	if cfg.border {
		cfg.style.border = []excelize.Border{
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
}

// Добавить заголовок
func xlsxAddTitle(xlsxFile *excelize.File, sheetName string) error {
	var err error
	text := cfg.title.text
	// Если text не указан, берем из A1 и строка 1 считается шапкой
	if text == "" {
		text, err = xlsxFile.GetCellValue(sheetName, "A1")
		if err != nil {
			return err
		}
	} else {
		// Иначе вставлеяем новую строку
		if err := xlsxFile.InsertRow(sheetName, 1); err != nil {
			return err
		}
	}

	// Добавляем текст
	if err := xlsxFile.SetCellStr(sheetName, "A1", text); err != nil {
		return err
	}
	// Font
	Font := excelize.Font{}
	Font.Bold = cfg.title.bold

	if len(cfg.title.color) == 6 {
		Font.Color = cfg.title.color
	}
	//Font.Family= "Times New Roman"
	if cfg.title.size != 0 {
		Font.Size = cfg.title.size
	}

	// Fill
	Fill := excelize.Fill{}

	if len(cfg.title.background) == 6 {
		Fill.Type = "pattern"
		Fill.Pattern = 1
		Fill.Color = append(Fill.Color, cfg.title.background)
	}
	// Alignment
	Alignment := cfg.style.alignment //параметры выравнивания как в документе
	Alignment.Horizontal = "center"

	// Создаем стиль
	Style, err := xlsxFile.NewStyle(&excelize.Style{
		Font:      &Font,
		Fill:      Fill,
		Border:    cfg.style.border, //параметры границы как в документе
		Alignment: &Alignment,
	})
	if err != nil {
		return err
	}

	// Применяем стиль
	if err := xlsxFile.SetCellStyle(sheetName, "A1", "A1", Style); err != nil {
		return err
	}

	// Объединем столбцы
	cols, err := xlsxFile.GetCols(sheetName)
	if err != nil {
		return err
	}

	lastColumn, err := excelize.ColumnNumberToName(len(cols))
	if err != nil {
		return err
	}

	if err := xlsxFile.MergeCell(sheetName, "A1", fmt.Sprintf("%s1", lastColumn)); err != nil {
		return err
	}

	return nil
}

// Обрабатываем параметры столбцов
func columnsWork(xlsxFile *excelize.File, sheetName string) error {
	cols, err := xlsxFile.GetCols(sheetName)
	if err != nil {
		return err
	}

	deleted := 0

	for _, column := range cfg.cols {
		// Удаляем столбцы
		if column.deleted {
			name, err := excelize.ColumnNumberToName(column.id)
			if err != nil {
				return err
			}
			if err := xlsxFile.RemoveCol(sheetName, name); err != nil {
				return err
			}
			deleted++
			continue
		}
		// Ширина столбцов
		switch column.width {
		case -1:
			break
		case 0:
			if err := colWidthAuto(xlsxFile, sheetName, column.id-deleted); err != nil {
				//return err
				fmt.Println("[WRN]\tcolWidthAuto: ", err.Error())
			}
		default:
			columnName, err := excelize.ColumnNumberToName(column.id - deleted)
			if err != nil {
				return err
			}
			if err := xlsxFile.SetColWidth(sheetName, columnName, columnName, float64(column.width)); err != nil {
				return err
			}
		}

		// Правила замены
		if len(column.replaces) > 0 {
			for _, replace := range column.replaces {
				// Перенос строки
				replaceTo := replace.to
				switch replace.to {
				case `\n`:
					replaceTo = "\n"
				case `\t`:
					replaceTo = "\t"
				}

				name, err := excelize.ColumnNumberToName(column.id - deleted)
				if err != nil {
					return err
				}
				col := cols[column.id-1]
				for n, rowCell := range col {
					if err := xlsxFile.SetCellStr(sheetName, fmt.Sprintf("%s%d", name, n+1), strings.ReplaceAll(rowCell, replace.from, replaceTo)); err != nil {
						return err
					}
				}
			}
		}

		// Правила поиска
		if len(column.finds) > 0 {
			for _, find := range column.finds {
				name, err := excelize.ColumnNumberToName(column.id - deleted)
				if err != nil {
					return err
				}
				col := cols[column.id-1]

				// Готовим общие настройки стилей
				findFont := excelize.Font{}
				findFill := excelize.Fill{}
				for _, action := range find.actions {
					switch action.name {
					case "bold":
						findFont.Bold = true
					case "size":
						size, err := strconv.Atoi(action.value)
						if err == nil {
							findFont.Size = float64(size)
						}
					case "color":
						if len(action.value) == 6 {
							findFont.Color = action.value
						}
					case "background":
						if len(action.value) == 6 {
							findFill.Type = "pattern"
							findFill.Pattern = 1
							findFill.Color = append(findFill.Color, action.value)
						}
					}
				}
				// Стиль для текущего find
				findStyle, err := xlsxFile.NewStyle(&excelize.Style{
					Font:      &findFont,
					Fill:      findFill,
					Alignment: &cfg.style.alignment,
					Border:    cfg.style.border,
				})
				if err != nil {
					return err
				}

				// перебираем строки
				for n, rowCell := range col {
					// Если в строке нашли текст
					if strings.Contains(rowCell, find.text) {

						// Если меняем стиль текста
						if find.target == "text" {
							// Разбиваем строку по найденому тексту
							ss := strings.Split(rowCell, find.text)
							var rtextall []excelize.RichTextRun // Общий итоговый текст ячейки

							// Отформатированный найденный текст
							var rfind excelize.RichTextRun
							rfind.Text = find.text
							rfind.Font = &findFont

							// Собираем итоговый текст
							for i := 0; i < len(ss); i++ {
								rtext := excelize.RichTextRun{
									Text: ss[i],
									Font: &excelize.Font{},
								}
								if i == len(ss)-1 {
									rtextall = append(rtextall, rtext)
								} else {
									rtextall = append(rtextall, rtext)
									rtextall = append(rtextall, rfind)
								}

							}

							// Заносим текст в ячейку
							if err := xlsxFile.SetCellRichText(sheetName, fmt.Sprintf("%s%d", name, n+1), rtextall); err != nil {
								return err
							}
						}

						// Если меняем стиль ячейки
						if find.target == "cell" {
							// Устанавливаем стиль ячейки
							if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", name, n+1), fmt.Sprintf("%s%d", name, n+1), findStyle); err != nil {
								return err
							}
						}

						// Если меняем стиль строки
						if find.target == "row" {
							// Устанавливаем стиль строки
							if err := xlsxFile.SetRowStyle(sheetName, n+1, n+1, findStyle); err != nil {
								return err
							}
						}
					}

				}
			}
		}

	}
	return nil
}

// задаем форматирование
func xlsxFormatSheet(xlsxFile *excelize.File, sheetName string) error {

	// Получаем данные о столбцах (могли измениться в columnsWork)
	cols, err := xlsxFile.GetCols(sheetName)
	if err != nil {
		return err
	}

	firstColumn, err := excelize.ColumnNumberToName(1)
	if err != nil {
		return err
	}
	lastColumn, err := excelize.ColumnNumberToName(len(cols))
	if err != nil {
		return err
	}

	// Создаем стиль всей таблицы
	wrapStyle, err := xlsxFile.NewStyle(&excelize.Style{
		//Font:      &sheetFont,
		//Fill:      sheetFill,
		Alignment: &cfg.style.alignment,
		Border:    cfg.style.border,
	})
	if err != nil {
		return err
	}

	// Применяем стиль все таблицы
	if err := xlsxFile.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s%d", lastColumn, len(cols[0])), wrapStyle); err != nil {
		return err
	}

	// Стиль заголовка
	if cfg.header.enable {
		if err := xlsxSetHeader(xlsxFile, sheetName, fmt.Sprintf("%s%d", firstColumn, cfg.header.row), fmt.Sprintf("%s%d", lastColumn, cfg.header.row)); err != nil {
			return err
		}

	}

	// Обрабатываем параметры столбцов
	if err := columnsWork(xlsxFile, sheetName); err != nil {
		return err
	}

	return nil
}

// Ширина столбца по содержимому
// colNum - номер столбца (начало с 1)
func colWidthAuto(xlsx *excelize.File, sheetName string, colNum int) error {
	cols, err := xlsx.GetCols(sheetName)
	if err != nil {
		return err
	}

	if len(cols) <= colNum-1 {
		return fmt.Errorf("[ERR] column %d not found, len: %d\n", colNum, len(cols))
	}

	col := cols[colNum-1]
	largestWidth := 0
	for r, rowCell := range col {
		// Если есть заголовок
		if r == 0 && colNum == 1 {
			continue
		}
		cellWidth := utf8.RuneCountInString(rowCell) + 2 // + 2 for margin
		if cellWidth > largestWidth {
			largestWidth = cellWidth
		}
	}
	name, err := excelize.ColumnNumberToName(colNum)
	if err != nil {
		return err
	}

	// Максимальная ширина 255 символов
	if largestWidth > 255 {
		largestWidth = 255
	}

	if err := xlsx.SetColWidth(sheetName, name, name, float64(largestWidth)); err != nil {
		return err
	}

	return nil
}

// Задает стиль заголовка
func xlsxSetHeader(xlsxFile *excelize.File, sheetName, startCell, endCell string) error {

	// Font
	headerFont := excelize.Font{}
	headerFont.Bold = cfg.header.bold

	if len(cfg.header.color) == 6 {
		headerFont.Color = cfg.header.color
	}
	//headerFont.Family= "Times New Roman"
	if cfg.header.size != 0 {
		headerFont.Size = cfg.header.size
	}

	// Fill
	headerFill := excelize.Fill{}

	if len(cfg.header.background) == 6 {
		headerFill.Type = "pattern"
		headerFill.Pattern = 1
		headerFill.Color = append(headerFill.Color, cfg.header.background)
	}

	// Создаем стиль заголовка
	headStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Font:      &headerFont,
		Fill:      headerFill,
		Border:    cfg.style.border,     //параметры границы как в документе
		Alignment: &cfg.style.alignment, //параметры выравнивания как в документе
	})
	if err != nil {
		return err
	}

	// Применяем стиль заголовка
	if err := xlsxFile.SetCellStyle(sheetName, startCell, endCell, headStyle); err != nil {
		return err
	}

	return nil
}
