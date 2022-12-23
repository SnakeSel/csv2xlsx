#!/bin/bash

## Colors
ClRed=$(tput setaf 1)
ClGreen=$(tput setaf 2)
ClYellow=$(tput setaf 3)
#ClBlue=$(tput setaf 4)
#ClMagenta=$(tput setaf 5)
#ClCyan=$(tput setaf 6)
#ClWhite=$(tput setaf 7)
Clreset=$(tput sgr0) #сброс цвета на стандартный
#Cltoend=$(tput hpa $(tput cols))$(tput cub 6) # сдвигает послед. текст до конца экрана

#####

press_and_cont(){
    echo ""
    read -n 1 -s -r -p "Нажмите любую кнопку для продолжения"
    echo ""
}


# Run go vet
if go vet ./...;then
    echo " - go vet ${ClGreen}OK${Clreset}"
else
    echo -e "\n\n - go vet ${ClRed}ERROR${Clreset}"
    press_and_cont
    exit 1
fi


# Run golangci-lint
if which -a golangci-lint >/dev/null 2>&1; then
    #linter="golangci-lint run -E gocritic -E stylecheck -E nestif -E revive -E govet"
    linter="golangci-lint run -E stylecheck -E revive -E govet"
    #linter="golangci-lint run"

    if $linter;then
        echo " - Linter ${ClGreen}OK${Clreset}"
    else
        echo -e "\n\n - Linter ${ClRed}ERROR${Clreset}"
        press_and_cont
        exit 1
    fi
else
    echo " - Linter ${ClYellow}skipped${Clreset}"
fi

# Build
version=$(date +%Y%m%d)

if go build -ldflags "-s -w -X 'main.version=${version}'";then
    echo " - Linux build ${ClGreen}OK${Clreset}"
fi

if GOOS=windows GOARCH=amd64 go build -ldflags "-s -w -X 'main.version=${version}'";then
    echo " - Windows build ${ClGreen}OK${Clreset}"
fi

press_and_cont

exit 0



