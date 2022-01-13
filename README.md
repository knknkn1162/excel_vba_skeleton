# README

Easy handling Excel VBA macros.

+ [x] support Japanese code
+ [x] Windows support
+ [x] MacOS support
+ [x] run from Powershell/bash

## OS features

||windows|MacOS|other(Linux)|
|---|---|---|---|
|export|o|o|o|
|import|o|o(\*1)|x|
|unbind macro|o|o(\*1)|x|
|run macro|o|o|x|

(\*1): use ./scripts/tools.bas(VBA macro). See also `Prerequisites > macOS` subsection.

## Prerequisites

### windows

+ Powershell >= 2.1
+ Install docker, chocolatey, nkf, make

```ps
# install nkf in Admin
Start-Process powershell -Verb runAs
choco install docker-for-windows
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf make
```

### macOS

+ Install docker, nkf, make
+ Create (empty) xlsm file and import ./scripts/tools.bas manually. Save as `helper.xlsm`.

# tutorial

Prepare xlsm file and set in `XLSM_FILE`. All you have to do is type `make <action> XLSM=${XLSM_FILE}`:

```sh
# unzip excelvba9.zip(in https://github.com/knknkn1162/excel_vba_playground/releases)
# import macro
make import XLSM=${XLSM_FILE}
# run macro(you can set ENTRYPOINT[default: main])
make run XLSM=${XLSM_FILE}
make export XLSM=${XLSM_FILE}
# remove macros in the xlsm VBProject
make unbind XLSM=${XLSM_FILE}
```

Note) If you run unimported macros, just `make import run XLSM=${XLSM_FILE}`.

## tools

```sh
# scaffold directory and macro according to ./scripts/template.bas
make template XLSM=${XLSM_FILE}
# git commit & push new macro, default: COMMIT_MSG=implement
make push XLSM=${XLSM_FILE}
```

# sample

+ [knknkn1162/vba100_knock](https://github.com/knknkn1162/vba100_knock)
