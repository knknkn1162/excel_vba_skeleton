SRC_ROOT_DIR=src
BOOKS_DIR=books
# used for nkf command
VBA_ENCODING=Shift_JIS
THIS_ENCODING=UTF-8
SRC_IMPORT_ROOT_DIR=$(SRC_ROOT_DIR)_$(VBA_ENCODING)
COMMIT_MSG=implement
VBAC_EXE=$(abspath ./vbac/vbac.wsf)
CAP_DIR=
ifneq (,$(CAP_DIR))
CAP_TARGET=clean-cap create-cap-dir
CAP_PATH=$(abspath $(CAP_DIR)/$(XLSM_BASENAME))
CAP_OPTION=/cap_dir:$(CAP_PATH)
endif

# `make <action> XLSM=ex008
XLSM_BASENAME=$(XLSM)
XLSM_NAME=$(XLSM).xlsm
XLSM_RELPATH=$(BOOKS_DIR)/$(XLSM_NAME)
XLSM_ABSPATH=$(abspath $(XLSM_RELPATH))
MACROS_DIR=$(abspath $(SRC_ROOT_DIR)/$(XLSM_BASENAME))
ENCODING_MACROS_DIR=$(abspath $(SRC_IMPORT_ROOT_DIR)/$(XLSM_BASENAME))
TARGETS=$(basename $(notdir $(wildcard $(BOOKS_DIR)/*.xlsm)))

ENTRYPOINT=main

# define macro
ifeq ("$(OS)", "Windows_NT")

define define-vbac-commands
$(1)-%: $(BOOKS_DIR)/%.xlsm
	make $(1) XLSM=$(notdir $$*)
$(1)-all: $(addprefix $(1)-, $(TARGETS))
endef

else
endif

# common commands

.PHONY: all imoprt export clean
all: export

COMMANDS=run \
		 import \
		 export \
		 unbind
## import commands
$(foreach cmd, $(COMMANDS), \
$(eval $(call define-vbac-commands,$(cmd))) \
)

template: create-xlsm-template
	-$(MKDIR) $(MACROS_DIR)
	cp ./templates/template.bas $(MACROS_DIR)/Module1.bas

push:
	make push -C $(SRC_ROOT_DIR) XLSM=$(XLSM) COMMIT_MSG="$(COMMIT_MSG)"
commit:
	make -C $(SRC_ROOT_DIR) XLSM=$(XLSM) COMMIT_MSG="$(COMMIT_MSG)"

pdftopng: $(patsubst %.pdf, %.png, $(wildcard $(CAP_PATH)/*.pdf))
%.png: %.pdf
	magick convert $^ -density 300 -quality 95 $@
run: run-macro
	make pdftopng



# OS dep. commands
ifeq ("$(OS)", "Windows_NT")
SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -Command
RM=rm -r -fo
TOUCH=New-Item -Type File
# see https://stackoverflow.com/a/47357220
MKDIR=mkdir -ea 0
ifeq (,$(wildcard $(SRC_ROOT_DIR)/))
endif
.PHONY: create-src-root-dir copy-import-dir
clean: clean-$(SRC_IMPORT_ROOT_DIR)
	$(RM) img

clean-$(SRC_IMPORT_ROOT_DIR):
	if ( Test-Path $(SRC_IMPORT_ROOT_DIR) ) { ${RM} $(SRC_IMPORT_ROOT_DIR) }
clean-cap:
	if ( Test-Path $(CAP_PATH) ) { ${RM} $(CAP_PATH) }
create-src-root-dir:
	if ( -not (Test-Path $(SRC_ROOT_DIR)) ) { mkdir $(SRC_ROOT_DIR) }
copy-import-dir: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
create-xlsm-template:
	if ( -not (Test-Path $(XLSM_ABSPATH))) { cp ./templates/empty.xlsm $(XLSM_ABSPATH) }
create-cap-dir:
	$(MKDIR) $(CAP_PATH)

# (try-)finally statement supports Ctrl-C in powershell. Whenever something error occurs in Excel Application, Ctrl-C can do cancellation and shutdown.
run-macro: $(CAP_TARGET)
	try { cscript $(VBAC_EXE) run /binary:$(XLSM_ABSPATH) /entrypoint:$(ENTRYPOINT) $(CAP_OPTION) } finally { Stop-Process -Name EXCEL }

export: clean-$(SRC_IMPORT_ROOT_DIR)
	if (-not ( Test-Path $(MACROS_DIR) )) { mkdir $(MACROS_DIR) }
	cscript $(VBAC_EXE) decombine /source:$(MACROS_DIR) /binary:$(XLSM_ABSPATH)
	Get-ChildItem -Recurse -Attributes !Directory $(MACROS_DIR)  | %{ nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) -Lu --overwrite $$_.FullName }

import: copy-import-dir
	Get-ChildItem -Recurse -Attributes !Directory $(ENCODING_MACROS_DIR)  | %{ nkf --ic=$(THIS_ENCODING) --oc=$(VBA_ENCODING) --overwrite $$_.FullName }
	cscript $(VBAC_EXE) combine /source:$(ENCODING_MACROS_DIR) /binary:$(XLSM_ABSPATH)

unbind:
	cscript $(VBAC_EXE) clear /binary:$(XLSM_ABSPATH)


# macOS or linux
else

RM=rm -rf
TOUCH=touch
MKDIR=mkdir -p
# Mac OS only
ifeq ("$(shell uname)", "Darwin")
run:
	./scripts/run_macro.scpt $(XLSM_ABSPATH) $(ENTRYPOINT)
else
run:
	$(error "run command is not implemented")
endif

HELPER_XLSM=$(abspath $(shell find $(CURDIR) -type f -name helper.xlsm))

clean-$(SRC_IMPORT_ROOT_DIR):
	$(RM) $(SRC_IMPORT_ROOT_DIR)
create-xlsm-template:
	cp --no-clobber ./templates/empty.xlsm $(XLSM_ABSPATH)

import: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
	find $(ENCODING_MACROS_DIR) -type f -print -exec nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) --overwrite {} \;
	./scripts/action_macos.scpt "import" $(HELPER_XLSM) $(XLSM_ABSPATH) $(ENCODING_MACROS_DIR)

unbind:
	./scripts/action_macos.scpt "unbind" $(HELPER_XLSM) $(XLSM_ABSPATH)

export:
	$(RM) $(MACROS_DIR)
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$(XLSM_RELPATH) --dst_dir /code/$(MACROS_DIR)

clean:
	$(RM) $(SRC_IMPORT_ROOT_DIR)
endif
