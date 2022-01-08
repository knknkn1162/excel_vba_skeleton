# README

Easy handle Excel VBA macros.

+ [x] support Japanese code
+ [x] Windows support
+ [x] MacOS support
+ [] Backup

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
+ Import ./scripts/tools.bas to
+ Create (empty) xlsm file, import ./scripts/tools.bas manually. Save as `helper.xlsm`.

# tutorial

Prepare xlsm file and set in `XLSM_FILE`. All you have to do is type `make <action> XLSM=${XLSM_FILE}`:

```sh
# unzip excelvba9.zip(in https://github.com/knknkn1162/excel_vba_playground/releases)
# import macro
make import XLSM=${XLSM_FILE}
make run XLSM=${XLSM_FILE}
# or you can set entrypoint(default: main)
make run XLSM=${XLSM_FILE} ENTRYPOINT=test
make export XLSM=${XLSM_FILE}
# remove macros in the xlsm VBProject
make unbind XLSM=${XLSM_FILE}
```

Note) If you run unimported macros, just `make import run XLSM=${XLSM_FILE}`.

## directories

```bash
./excelvba1
├── ex001.xlsm
└── ex007.xlsm
./src/excelvba1
├── ex001.xlsm
│   └─ Module1.bas
└── ex007.xlsm
    └── Module1.bas
```
