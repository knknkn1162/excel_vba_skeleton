# README

Easy handling Excel VBA project.

+ [x] support Japanese code

## OS features

||windows|MacOS|other|
|---|---|---|---|
|export|o|o|o|
|import|o|x|x|
|unbind macro|o|x|x|
|run macro|o|o|x|

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

## how to run excel macro

0. Fetch this repo; `git clone https://github.com/knknkn1162/excel_vba_playground`

1. Download xlsm books in https://github.com/knknkn1162/excel_vba_playground/releases not containing VBA macro.

2. Type commands below;

```sh
unzip excelvba8.zip
# import macro
make import
# make run-<dir>-<.xlsm file>
make run-vba100-ex001.xlsm
# run macros in directory
make run-vba100 # or make run

# in mac OS
make run XLSM=<path_to_xlsm>
```

3. If you edit macro source, export it to text format automatically!

```sh
make export
```

Note) If you clean macro source in your .xlsm, try `make unbind`:

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
