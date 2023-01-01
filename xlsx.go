// xlsx.go содержит функции манипуляций с готовым файлом xlsx
//

package main

import (
	"fmt"

	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Задать общий стиль
func xlsxSetDefaultStyle(style *_style, border bool) {
	// выравнивание текста
	style.Alignment.WrapText = true

	//style.Font.Family = "Times New Roman"

	// Добавить границу
	if border {
		style.Border = []excelize.Border{
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
	// Font.Family= "Times New Roman"
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
	Alignment := cfg.style.Alignment //параметры выравнивания как в документе
	Alignment.Horizontal = "center"

	// Создаем стиль
	Style, err := xlsxFile.NewStyle(&excelize.Style{
		Font:      &Font,
		Fill:      Fill,
		Border:    cfg.style.Border, //параметры границы как в документе
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
			name, err := excelize.ColumnNumberToName(column.id - deleted)
			if err != nil {
				return err
			}
			if err := xlsxFile.RemoveCol(sheetName, name); err != nil {
				return err
			}
			deleted++
			continue
		}

		// Получаем имя текущего столбца
		columnName, err := excelize.ColumnNumberToName(column.id - deleted)
		if err != nil {
			return err
		}

		// Ширина столбцов
		switch column.width {
		case -1:
			// Пропускаем
		case 0:
			if err := colWidthAuto(xlsxFile, sheetName, column.id-deleted); err != nil {
				//return err
				fmt.Println("[WRN]\tcolWidthAuto: ", err.Error())
			}
		default:
			if err := xlsxFile.SetColWidth(sheetName, columnName, columnName, float64(column.width)); err != nil {
				return err
			}
		}

		// Базовый стиль столбца
		colStyleIsDefault := true
		colStyleDefault := _style{
			Font:      cfg.style.Font,
			Fill:      cfg.style.Fill,
			Border:    cfg.style.Border,
			Alignment: cfg.style.Alignment,
		}

		if column.horizontal != "" {
			colStyleIsDefault = false
			if err := setAlignment(&colStyleDefault.Alignment, "horizontal", column.horizontal); err != nil {
				fmt.Printf("[WRN]\tcolumnsWork[%d]: %s\n", column.id, err.Error())
				break
			}

		}
		// Стиль для текущего column
		colStyleID, err := xlsxFile.NewStyle(&excelize.Style{
			Alignment: &colStyleDefault.Alignment,
			Font:      &colStyleDefault.Font,
			Fill:      colStyleDefault.Fill,
			Border:    colStyleDefault.Border,
		})
		if err != nil {

			fmt.Printf("[WRN]\tcolumnsWork[%d]: %s\n", column.id, err.Error())
			break
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

				rows := cols[column.id-1]
				for n, rowCell := range rows {
					if err := xlsxFile.SetCellStr(sheetName, fmt.Sprintf("%s%d", columnName, n+1), strings.ReplaceAll(rowCell, replace.from, replaceTo)); err != nil {
						return err
					}
				}
			}
		}

		// Правила поиска
		if len(column.finds) > 0 || !colStyleIsDefault {

			// Заполняем стили для всех finds
			for findId, find := range column.finds {
				column.finds[findId].style = newStyleFind(find, colStyleDefault)
			}

			// перебираем строки
			rows := cols[column.id-1]
			for n, rowCell := range rows {

				// Применяем стиль столбца (если менялся)
				if !colStyleIsDefault {
					// Устанавливаем стиль ячейки
					if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", columnName, n+1), fmt.Sprintf("%s%d", columnName, n+1), colStyleID); err != nil {
						return err
					}
				}

				for _, find := range column.finds {

					// Если в строке нашли текст
					if strings.Contains(rowCell, find.text) {

						// Создаем стиль документа
						findStyleID, err := xlsxFile.NewStyle(&excelize.Style{
							Font:      &find.style.Font,
							Fill:      find.style.Fill,
							Border:    find.style.Border,
							Alignment: &find.style.Alignment,
						})
						if err != nil {
							return err
						}

						switch find.target {
						case "text": // Если меняем стиль текста
							// Получаем отформатированный текст
							rtextall := getFindRichText(rowCell, find.text, &find.style.Font)

							// Заносим текст в ячейку
							if err := xlsxFile.SetCellRichText(sheetName, fmt.Sprintf("%s%d", columnName, n+1), rtextall); err != nil {
								return err
							}
						case "cell": // Если меняем стиль ячейки
							// Устанавливаем стиль ячейки
							if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", columnName, n+1), fmt.Sprintf("%s%d", columnName, n+1), findStyleID); err != nil {
								return err
							}
						case "row": // Если меняем стиль строки

							// Устанавливаем стиль строки
							if err := xlsxFile.SetRowStyle(sheetName, n+1, n+1, findStyleID); err != nil {
								return err
							}

						}
					}
				} // for finds

			} // for range col

		} // if column.finds

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

	lastColumn, err := excelize.ColumnNumberToName(len(cols))
	if err != nil {
		return err
	}

	// Создаем стиль всей таблицы

	wrapStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Font: &cfg.style.Font,
		//Fill:      sheetFill,
		Alignment: &cfg.style.Alignment,
		Border:    cfg.style.Border,
	})
	if err != nil {
		return err
	}

	// Применяем стиль все таблицы
	if err := xlsxFile.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s%d", lastColumn, len(cols[0])), wrapStyle); err != nil {
		return err
	}

	// Обрабатываем параметры столбцов
	if err = columnsWork(xlsxFile, sheetName); err != nil {
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
		return fmt.Errorf("[ERR] column %d not found, len: %d", colNum, len(cols))
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

	// Alignment
	headerAlignment := cfg.style.Alignment //параметры выравнивания как в документе

	if err := setAlignment(&headerAlignment, "horizontal", cfg.header.horizontal); err != nil {
		fmt.Printf("[WRN]\txlsxSetHeader: %s\n", err.Error())
	}

	// Создаем стиль заголовка
	headStyle, err := xlsxFile.NewStyle(&excelize.Style{
		Font:      &headerFont,
		Fill:      headerFill,
		Border:    cfg.style.Border, //параметры границы как в документе
		Alignment: &headerAlignment,
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

// Задать парметры выравнивания
// param = horizontal|vertical
func setAlignment(alignment *excelize.Alignment, param, value string) error {

	switch param {
	case "horizontal":
		switch value {
		case "left", "center", "right", "fill", "distributed", "justify", "centerContinuous":
			alignment.Horizontal = value
		case "":
			break
		default:
			return fmt.Errorf("unknown value: %s", value)
		}
	case "vertical":
		switch value {
		case "top", "center", "justify", "distributed":
			alignment.Vertical = value
		case "":
			break
		default:
			fmt.Println("unknown value: ", value)
		}
	default:
		return fmt.Errorf("unknown param: %s", param)
	}

	return nil
}

// getFindRichText Получить отформатированный текст
// row: полный текст строки
// find: текст который необходимо отформатировать
// font: font для отформатированного текста
func getFindRichText(row, find string, font *excelize.Font) []excelize.RichTextRun {
	// Разбиваем строку по найденому тексту
	ss := strings.Split(row, find)
	var rtextall []excelize.RichTextRun // Общий итоговый текст ячейки

	// Отформатированный найденный текст
	var rfind excelize.RichTextRun
	rfind.Text = find
	rfind.Font = font

	// Собираем итоговый текст
	// Идем по всем элементам разбитого текста.
	// Между элементами - наш отформатированный текст
	for i := 0; i < len(ss); i++ {
		// rtext - RichTextRun для текущего элемента
		rtext := excelize.RichTextRun{
			Text: ss[i],
			Font: &excelize.Font{},
		}
		// Если последний элемент, то дабавляем только его
		if i == len(ss)-1 {
			rtextall = append(rtextall, rtext)
		} else { // Иначе элемент + отформатированный текст
			rtextall = append(rtextall, rtext)
			rtextall = append(rtextall, rfind)
		}

	}

	return rtextall
}

// Создаем стиль для find
func newStyleFind(find _find, defStyle _style) _style {
	findStyle := defStyle

	for _, action := range find.actions {
		switch action.name {
		case "bold":
			findStyle.Font.Bold = true
		case "size":
			size, err := strconv.Atoi(action.value)
			if err == nil {
				findStyle.Font.Size = float64(size)
			} else {
				fmt.Printf("[WRN]\tcolumnsWork|find|%s: unknown size: %s\n", find.text, action.value)
			}
		case "color":
			if len(action.value) == 6 {
				findStyle.Font.Color = action.value
			} else {
				fmt.Printf("[WRN]\tcolumnsWork|find|%s: unknown color: %s\n", find.text, action.value)
			}
		case "background":
			if len(action.value) == 6 {
				findStyle.Fill.Type = "pattern"
				findStyle.Fill.Pattern = 1
				findStyle.Fill.Color = append(findStyle.Fill.Color, action.value)
			} else {
				fmt.Printf("[WRN]\tcolumnsWork|find|%s: unknown background: %s\n", find.text, action.value)
			}
		case "horizontal":
			if err := setAlignment(&findStyle.Alignment, "horizontal", action.value); err != nil {
				fmt.Printf("[WRN]\tcolumnsWork|find|%s: %s\n", find.text, err.Error())
			}

		case "vertical":
			if err := setAlignment(&findStyle.Alignment, "vertical", action.value); err != nil {
				fmt.Printf("[WRN]\tcolumnsWork|find|%s: %s\n", find.text, err.Error())
			}

		default:
			fmt.Printf("[WRN]\tcolumnsWork|find|%s: unknown action: %s\n", find.text, action.name)
		} // switch
	} // for range find.actions

	return findStyle
}
