// xlsx.go содержит функции манипуляций с готовым файлом xlsx
//

//nolint:cyclop // TODO: вынести в отдельный модуль
package main

import (
	"fmt"

	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Добавить заголовок
func xlsxAddTitle(xlsxFile *excelize.File, sheetName string, title _title, style _style) error {

	var err error

	// Если text не указан, берем из A1 и строка 1 считается шапкой
	if title.text == "" {
		title.text, err = xlsxFile.GetCellValue(sheetName, "A1")
		if err != nil {
			return err
		}
	} else {
		// Иначе вставлеяем новую строку
		if err := xlsxFile.InsertRows(sheetName, 1, 1); err != nil {
			return err
		}
	}

	// Добавляем текст
	if err := xlsxFile.SetCellStr(sheetName, "A1", title.text); err != nil {
		return err
	}

	// Формируем стиль
	style.Font.Bold = title.bold

	if len(title.color) == 6 {
		style.Font.Color = title.color
	}
	// Font.Family= "Times New Roman"
	if title.size != 0 {
		style.Font.Size = title.size
	}

	if len(title.background) == 6 {
		style.Fill.Type = "pattern"
		style.Fill.Pattern = 1
		style.Fill.Color = append(style.Fill.Color, title.background)
	}
	// Alignment
	style.Alignment.Horizontal = "center"
	style.Alignment.Vertical = "center"

	// Создаем стиль
	styleID, err := xlsxNewStyleID(xlsxFile, style)
	if err != nil {
		return err
	}

	// Применяем стиль
	if err := xlsxFile.SetCellStyle(sheetName, "A1", "A1", styleID); err != nil {
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
func xlsxSetColumnFormat(xlsxFile *excelize.File, sheetName string, colsParam []_column, style _style) error {
	cols, err := xlsxFile.GetCols(sheetName)
	if err != nil {
		return err
	}

	deleted := 0

	for _, column := range colsParam {
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
				// return err
				fmt.Println("[WRN]\tcolWidthAuto: ", err.Error())
			}
		default:
			if err := xlsxFile.SetColWidth(sheetName, columnName, columnName, float64(column.width)); err != nil {
				return err
			}
		}

		// Стиль столбца из общего стиля
		colStyleDefault, colStyleIsChanged := xlsxGetColumnStyleFromSettings(style, column)

		// StyleID для текущего column
		colStyleID, err := xlsxNewStyleID(xlsxFile, colStyleDefault)
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

		// Правила поиска или изменение стиля
		if len(column.finds) > 0 || colStyleIsChanged {
			// Заполняем стили для всех finds
			for findID, find := range column.finds {
				column.finds[findID].style = newStyleFind(find, colStyleDefault)
				column.finds[findID].styleID, err = xlsxNewStyleID(xlsxFile, column.finds[findID].style)
				if err != nil {
					fmt.Println(err.Error())
					break
				}
			}

			if err := colApplyFind(xlsxFile, sheetName, columnName, column, cols[column.id-1], colStyleIsChanged, colStyleID); err != nil {
				return err
			}
		}

	}

	return nil
}

func colApplyFind(xlsxFile *excelize.File, sheetName string, columnName string, column _column, colRows []string, colStyleIsChanged bool, colStyleID int) error {

	// перебираем строки
	for n, rowCell := range colRows {

		// Применяем стиль столбца (если менялся)
		if colStyleIsChanged {
			// Устанавливаем стиль ячейки
			if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", columnName, n+1), fmt.Sprintf("%s%d", columnName, n+1), colStyleID); err != nil {
				return err
			}
		}

		for _, find := range column.finds {

			// Если в строке нашли текст
			if strings.Contains(rowCell, find.text) {

				switch find.target {
				case "text": // Если меняем стиль текста
					// Получаем отформатированный текст
					//nolint:gosec // TODO: проверить!!!!
					rtextall := getFindRichText(rowCell, find.text, &find.style.Font)

					// Заносим текст в ячейку
					if err := xlsxFile.SetCellRichText(sheetName, fmt.Sprintf("%s%d", columnName, n+1), rtextall); err != nil {
						return err
					}
				case "cell": // Если меняем стиль ячейки
					// Устанавливаем стиль ячейки
					if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", columnName, n+1), fmt.Sprintf("%s%d", columnName, n+1), find.styleID); err != nil {
						return err
					}
				case "row": // Если меняем стиль строки

					// Устанавливаем стиль строки
					if err := xlsxFile.SetRowStyle(sheetName, n+1, n+1, find.styleID); err != nil {
						return err
					}

				}
			}
		}

	}

	return nil
}

// задаем форматирование всей таблице
func xlsxSetTableStyle(xlsxFile *excelize.File, sheetName string, style _style, defColumn _defColParam) error {

	// Получаем данные о столбцах
	cols, err := xlsxFile.GetCols(sheetName)
	if err != nil {
		return err
	}

	lastColumn, err := excelize.ColumnNumberToName(len(cols))
	if err != nil {
		return err
	}

	for id := 1; id <= len(cols); id++ {
		// Ширина столбцов
		switch defColumn.width {
		case -1:
			// Пропускаем
		case 0:
			if err := colWidthAuto(xlsxFile, sheetName, id); err != nil {
				// return err
				fmt.Println("[WRN]\tcolWidthAuto: ", err.Error())
			}
		default:
			columnName, err := excelize.ColumnNumberToName(id)
			if err != nil {
				return err
			}
			if err := xlsxFile.SetColWidth(sheetName, columnName, columnName, float64(defColumn.width)); err != nil {
				return err
			}
		}
	}

	// Создаем стиль всей таблицы
	wrapStyle, err := xlsxNewStyleID(xlsxFile, style)
	if err != nil {
		return err
	}

	// Применяем стиль все таблицы
	if err := xlsxFile.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s%d", lastColumn, len(cols[0])), wrapStyle); err != nil {
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
func xlsxSetHeader(xlsxFile *excelize.File, sheetName, startCell, endCell string, header _header, style _style) error {

	// Font
	style.Font.Bold = header.bold

	if len(header.color) == 6 {
		style.Font.Color = header.color
	}
	// headerFont.Family= "Times New Roman"
	if header.size != 0 {
		style.Font.Size = header.size
	}

	// Fill
	if len(header.background) == 6 {
		style.Fill.Type = "pattern"
		style.Fill.Pattern = 1
		style.Fill.Color = append(style.Fill.Color, header.background)
	}

	// Alignment
	if err := setAlignment(&style.Alignment, "horizontal", header.horizontal); err != nil {
		fmt.Printf("[WRN]\txlsxSetHeader: %s\n", err.Error())
	}

	// Создаем стиль заголовка
	headStyleID, err := xlsxNewStyleID(xlsxFile, style)
	if err != nil {
		return err
	}

	// Применяем стиль заголовка
	if err := xlsxFile.SetCellStyle(sheetName, startCell, endCell, headStyleID); err != nil {
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

// Создаем новый стиль документа
func xlsxNewStyleID(xlsxFile *excelize.File, style _style) (int, error) {

	return xlsxFile.NewStyle(&excelize.Style{
		Alignment: &style.Alignment,
		Font:      &style.Font,
		Fill:      style.Fill,
		Border:    style.Border,
	})

}

// Стиль столбца из общего стиля
func xlsxGetColumnStyleFromSettings(style _style, column _column) (_style, bool) {
	styleIsChanged := false
	colStyle := style

	if column.horizontal != "" {
		styleIsChanged = true
		if err := setAlignment(&colStyle.Alignment, "horizontal", column.horizontal); err != nil {
			fmt.Printf("[WRN]\txlsxGetColumnStyleFromSettings[%d]: %s\n", column.id, err.Error())
		}

	}
	if column.size != 0 {
		styleIsChanged = true
		colStyle.Font.Size = column.size
	}

	return colStyle, styleIsChanged
}
