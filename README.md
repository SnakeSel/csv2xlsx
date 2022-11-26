# csv2xlsx

Ковертирование csv в xlsx с форматированием.

```
csv2xlsx [command] <args>
command:
  -c string
    Path to the config file (optional)
  -d string
    Delimiter for felds in the CSV input (optional) (default "\t")
  -f string
    Path to the CSV input file
  -o string
    Path to the XLSX output file
```
## Использвание
#### Простая конвертация csv в xlsx
```sh
csv2xlsx -f input.csv -o output.xlsx
```
#### Используем файл настроек
```sh
csv2xlsx -c config.cfg -f input.csv -o output.xlsx
```
## Файл настроек
##### Параметры вне секции ini файла:
 - delimiter = разделитель столбцов в csv (по умолчанию \t)
 - sheet = имя листа (по умолчанию Sheet1)
 - title = Добавить шапку с текстом
 - border = рисовать границу.(логическое)
 - header = выделять первую строку заголовком (логическое)

##### [Имя секции] - номер столбца (начало с 1)
 - width = ширина столбца (0 ширина по тексту)
 - replace = "строка для замены","на что заменяем"
 - delete = удалить столбец из xlsx (логическое)
 - find = "text",action,action,...

##### action:
Применяются к найденому тексту:
  - size = число. Размер текста.
  - color = строка, 6 символов. Цвет (пример: FF0000)
  - bold Наличие параметра делает текст жирным

Применяются ко всей ячейке с найденным текстом:
  - cellsize = 12
  - cellcolor = FFA500
  - cellbold
  - cellbackground = FFA500

##### Пример файла настроек:
```
# delimiter = разделитель столбцов в csv (по умолчанию \t)
delimiter="\t"

#sheet= имя листа (по умолчанию Sheet1)
sheet=Лист1

# title = Добавить шапку с текстом
title="CSV to XLSX"

# border = рисовать границу.(логическое)
border=1

# header = выделять первую строку заголовком (логическое)
header=1

[1]
width=0

[2]
width=0
find="RED",color=FF0000,size=12,bold
find=Внимание,cellbold,cellbackground=FFA500,cellsize=14

[3]
width=100
replace = "/ ","\n"
replace = "to","2"
```

## Сборка из исходников
#### Загружаем исходный код:
```sh
$ git clone https://github.com/SnakeSel/csv2xlsx
```
#### Переходим в директорию проекта:
```sh
$ cd csv2xlsx
```
#### Компилируем:
```sh
go build
```
