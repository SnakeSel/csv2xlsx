#!/bin/bash

binName="csv2xlsx"
version=$(date +%Y%m%d)

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

# входные прараметры - список доп файлов для добавления в архив
create_arhive(){
    mkdir -p "./build"

    if tar -caf "./build/${binName}-${version}_linux.tar.gz" ${binName} "$@";then
        echo " - Linux archive create ${ClGreen}OK${Clreset}"
    fi

    if [ -f "${binName}.exe" ];then
        if zip -r "./build/${binName}-${version}_windows.zip" ${binName}.exe "$@";then
            echo " - Windows archive create ${ClGreen}OK${Clreset}"
        fi
    fi
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
    linter="golangci-lint run"

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
if go build -ldflags "-s -w -X 'main.version=${version}'" -o "${binName}";then
    echo " - Linux build ${ClGreen}OK${Clreset}"
fi

if GOOS=windows GOARCH=amd64 go build -ldflags "-s -w -X 'main.version=${version}'" -o "${binName}.exe";then
    echo " - Windows build ${ClGreen}OK${Clreset}"
fi

chmod +x csv2xlsx

# Create archive
create_arhive "example.cfg"

press_and_cont

exit 0
