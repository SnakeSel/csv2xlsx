#!/bin/bash

linter="golangci-lint run -E gocritic -E stylecheck -E nestif -E revive -E govet"
linter="golangci-lint run -E stylecheck -E revive -E govet"
#linter="golangci-lint run"

if $linter ;then
    echo "linter OK"
    if go build -ldflags "-s -w -X 'main.version=$(date +%Y%m%d)'";then
        echo "build OK"
    fi
fi

read -n 1 -s -r -p "Нажмите любую кнопку для продолжения"
echo""

exit 0



