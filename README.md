# csv2xlsx

Конвертирование csv в xlsx с форматированием.

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
 - delimiter = Разделитель столбцов в csv (по умолчанию \t);
 - sheet = Имя листа (по умолчанию Sheet1);
 - border = Рисовать границу (1|0);
 - horizontal = Выравнивание текста в ячейках по горизонтали (left|center|right|fill|distributed);
 - vertical = Выравнивание текста в ячейках по вертикали (top|center|justify|distributed);

#### [title] - Добавить шапку с текстом
 - enable = Добавлять шапку (1|0);
 - text = Текст шапки (Если пусто - встанет текст из A1 и строка 1 считается шапкой; Если задан - вставляем новую строку с текстом);
 - size = Размер текста (число);
 - bold = Жирный текст (1|0);
 - color = Цвет текста (строка, 6 символов)(пример: FF0000);
 - background = Цвет заливки (строка, 6 символов)(пример: FFA500);

#### [header] - Выделять строку заголовком
 - enable = Выделять строку заголовком (1|0);
 - row = Номер строки. Соответствует исходному файлу (число)(по умолчанию 1);
 - size = Размер текста (число);
 - bold = Жирный текст (1|0);
 - color = Цвет текста (строка, 6 символов)(пример: FF0000);
 - background = Цвет заливки (строка, 6 символов)(пример: FFA500);

##### [Имя секции] - номер столбца (начало с 1)
 - width = ширина столбца (0 ширина по тексту)
 - replace = "строка для замены","на что заменяем"
 - delete = удалить столбец из xlsx (логическое)
 - find = text,target,action,action,...
##### target (для find):
 - text Применяются к найденому тексту
 - cell Применяются ко всей ячейке с найденным текстом
 - row Применяются ко всей строке с найденным текстом
##### action (для find):
  - size = Размер текста (число);
  - color = Цвет текста (строка, 6 символов)(пример: FF0000);
  - bold Наличие параметра делает текст жирным
  - background = Цвет заливки (строка, 6 символов)(пример: FFA500) Не работает с target=text;
  - horizontal = Выравнивание текста по горизонтали (left|center|right|fill|distributed);
  - vertical = Выравнивание текста по вертикали (top|center|justify|distributed);

### Пример файла настроек:
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
find="RED",text,color=FF0000,size=12,bold
find=Внимание,row,bold,background=FFA500,size=14

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
