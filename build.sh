#!/bin/bash

## Colors
#ClRed=$(tput setaf 1)
ClGreen=$(tput setaf 2)
#ClYellow=$(tput setaf 3)
#ClBlue=$(tput setaf 4)
#ClMagenta=$(tput setaf 5)
#ClCyan=$(tput setaf 6)
#ClWhite=$(tput setaf 7)
Clreset=$(tput sgr0) #сброс цвета на стандартный
#Cltoend=$(tput hpa $(tput cols))$(tput cub 6) # сдвигает послед. текст до конца экрана

#####

#linter="golangci-lint run -E gocritic -E stylecheck -E nestif -E revive -E govet"
linter="golangci-lint run -E stylecheck -E revive -E govet"
#linter="golangci-lint run"

if ! $linter;then
    echo ""
    read -n 1 -s -r -p "Нажмите любую кнопку для продолжения"
    echo ""
    exit 1
fi

echo " - Linter ${ClGreen}OK${Clreset}"

version=$(date +%Y%m%d)

if go build -ldflags "-s -w -X 'main.version=${version}'";then
    echo " - Linux build ${ClGreen}OK${Clreset}"
fi

if GOOS=windows GOARCH=amd64 go build -ldflags "-s -w -X 'main.version=${version}'";then
    echo " - Windows build ${ClGreen}OK${Clreset}"
fi

echo ""
read -n 1 -s -r -p "Нажмите любую кнопку для продолжения"
echo ""

exit 0



