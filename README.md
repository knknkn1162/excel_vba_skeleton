# README

Easy handling Excel VBA project.

+ [x] support Japanese code

## OS features

||windows|MacOS|other|
|---|---|---|---|
|export|o|o|o|
|import|o|o(\*1)|x|
|unbind macro|o|o(\*1)|x|
|run macro|o|o|x|

(\*1): use ./scripts/tools.bas(VBA macro).

## Prerequisites

### windows

+ install docker, nkf

```ps
# install nkf in Admin
Start-Process powershell -Verb runAs
choco install docker-for-windows
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf
```

### mac

+ install docker, nkf
+ import ./scripts/tools.bas to (empty) xlsm file manually and save as `helper.xlsm`.

## how to run excel macro

0. Fetch this repo; `git clone https://github.com/knknkn1162/excel_vba_playground`

1. Download xlsm books in https://github.com/knknkn1162/excel_vba_playground/releases not containing VBA macro.

2. Type commands below;

```sh
unzip excelvba8.zip
# import macro
make import XLSM=${XLSM_FILE}
make run XLSM=${XLSM_FILE}
```

Note) if you exec all files in directory(such as DirA), type `make import-DirA`.

3. If you edit macro source, export it to text format automatically!

```sh
make export XLSM=${XLSM_FILE}
```

Note) If you clean macro source in your .xlsm, try `make unbind`:

```sh
make unbind XLSM=${XLSM_FILE}
```

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
