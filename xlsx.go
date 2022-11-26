package main

import (
	"fmt"

	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Добавить заголовок
func xlsxAddTitle(xlsxFile *excelize.File, sheetName string, title string) error {
	if err := xlsxFile.InsertRow(sheetName, 1); err != nil {
		return err
	}

	if err := xlsxFile.SetCellRichText(sheetName, "A1", []excelize.RichTextRun{
		{
			Text: title,
			Font: &excelize.Font{
				Bold: true,
				Size: 16,
				//Color:  "2354e8",
				//Family: "Times New Roman",
			},
		},
	}); err != nil {
		return err
	}

	style, err := xlsxFile.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return err
	}

	if err := xlsxFile.SetCellStyle(sheetName, "A1", "A1", style); err != nil {
		return err
	}

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
				return err
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
							rfind.Font = &excelize.Font{}
							for _, action := range find.actions {
								switch action.name {
								case "bold":
									rfind.Font.Bold = true
								case "size":
									size, err := strconv.Atoi(action.value)
									if err == nil {
										rfind.Font.Size = float64(size)
									}
								case "color":
									if len(action.value) == 6 {
										rfind.Font.Color = action.value
									}
								}

							}

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

							// orStyle, err := xlsxFile.GetCellStyle(sheetName, fmt.Sprintf("%s%d", name, n+1))
							// if err != nil {
							// 	return err
							// }
							cellFont := excelize.Font{}
							cellFill := excelize.Fill{}

							for _, action := range find.actions {
								switch action.name {
								case "bold":
									cellFont.Bold = true
								case "size":
									size, err := strconv.Atoi(action.value)
									if err == nil {
										cellFont.Size = float64(size)
									}
								case "color":
									if len(action.value) == 6 {
										cellFont.Color = action.value
									}
								case "background":
									if len(action.value) == 6 {
										cellFill.Type = "pattern"
										cellFill.Pattern = 1
										cellFill.Color = append(cellFill.Color, action.value)
									}
								}
							}

							cellStyle, err := xlsxFile.NewStyle(&excelize.Style{
								Font: &cellFont,
								Fill: cellFill,
							})
							if err != nil {
								return err
							}

							// Устанавливаем стиль ячейки
							if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s%d", name, n+1), fmt.Sprintf("%s%d", name, n+1), cellStyle); err != nil {
								return err
							}
						}

						// Если меняем стиль строки
						if find.target == "row" {

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

	//sheetFont := excelize.Font{}
	//sheetFill := excelize.Fill{}
	sheetAlignment := excelize.Alignment{}
	sheetBorder := []excelize.Border{}

	// Формат текста
	sheetAlignment.WrapText = true
	sheetAlignment.Vertical = "center"

	// Добавить границу
	if cfg.border {
		sheetBorder = []excelize.Border{
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

	// Создаем стиль всей таблицы
	wrapStyle, err := xlsxFile.NewStyle(&excelize.Style{
		//Font:      &sheetFont,
		//Fill:      sheetFill,
		Alignment: &sheetAlignment,
		Border:    sheetBorder,
	})
	if err != nil {
		return err
	}

	// Применяем стиль все таблицы
	if err := xlsxFile.SetCellStyle(sheetName, "A1", fmt.Sprintf("%s%d", lastColumn, len(cols[0])), wrapStyle); err != nil {
		return err
	}

	// Стиль заголовка
	if cfg.header {
		headerFont := excelize.Font{}

		headerFont.Bold = true
		//headerFont.Color=  "2354e8"
		//headerFont.Family= "Times New Roman"
		headerFont.Size = 14

		// Создаем стиль заголовка
		headStyle, err := xlsxFile.NewStyle(&excelize.Style{
			Font:   &headerFont,
			Border: sheetBorder, //параметры границы как в документе
		})
		if err != nil {
			return err
		}
		// Применяем стиль заголовка
		if err := xlsxFile.SetCellStyle(sheetName, fmt.Sprintf("%s1", firstColumn), fmt.Sprintf("%s1", lastColumn), headStyle); err != nil {
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

	col := cols[colNum-1]
	largestWidth := 0
	for _, rowCell := range col {
		cellWidth := utf8.RuneCountInString(rowCell) + 2 // + 2 for margin
		if cellWidth > largestWidth {
			largestWidth = cellWidth
		}
	}
	name, err := excelize.ColumnNumberToName(colNum)
	if err != nil {
		return err
	}

	if err := xlsx.SetColWidth(sheetName, name, name, float64(largestWidth)); err != nil {
		return err
	}

	return nil
}
