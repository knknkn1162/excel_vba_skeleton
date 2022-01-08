DIRS=excelvba9 \
	 excelvba1 \
	 vba100

XLSMS=$(filter-out $(wildcard */~*.xlsm), \
	  $(foreach dir, $(DIRS), $(wildcard $(dir)/*.xlsm)) \
)
TARGETS=$(basename $(XLSMS))
SRC_ROOT_DIR=src
# used for nkf command
VBA_ENCODING=Shift_JIS
THIS_ENCODING=UTF-8
SRC_IMPORT_ROOT_DIR=$(SRC_ROOT_DIR)_$(VBA_ENCODING)
COMMIT_MSG=implement
VBAC_EXE=$(abspath ./vbac/vbac.wsf)

XLSM_ABSPATH=$(abspath $(XLSM))
XLSM_NAME=$(notdir $(XLSM))
XLSM_PARENT_DIR=$(lastword $(subst /, ,$(dir $(abspath $(XLSM)))))
XLSM_RELPATH=$(XLSM_PARENT_DIR)/$(XLSM_NAME)
MACROS_DIR=$(SRC_ROOT_DIR)/$(XLSM_RELPATH)

ENTRYPOINT=main

# define macro
ifeq ("$(OS)", "Windows_NT")

define define-vbac-commands
$(2)-$(1)-%: $(1)/%
	make $(2) XLSM=$$^
$(2)-$(1): $(1) $(addprefix $(2)-$(1)-, $(notdir $(wildcard $(1)/*.xlsm)))
endef

else
endif

# common commands

.PHONY: all imoprt export clean
all: export

## run commands
$(foreach dir, $(DIRS), \
$(eval $(call define-vbac-commands,$(dir), run)) \
)

## import commands
$(foreach dir, $(DIRS), \
$(eval $(call define-vbac-commands,$(dir), import)) \
)

## export commands
$(foreach dir, $(DIRS), \
$(eval $(call define-vbac-commands,$(dir), export)) \
)

## unbind commands
$(foreach dir, $(DIRS), \
$(eval $(call define-vbac-commands,$(dir), unbind)) \
)

touch_module:
	mkdir $(MACROS_DIR)
	$(TOUCH) $(MACROS_DIR)/Module.bas

push:
	git push
commit:
	git add $(MACROS_DIR)
	git commit -m "implement $(XLSM_RELPATH)"

# OS dep. commands
ifeq ("$(OS)", "Windows_NT")
SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -Command
RM=rm -r -fo
TOUCH=New-Item -Type File
ifeq (,$(wildcard $(SRC_ROOT_DIR)/))
endif
.PHONY: create-src-root-dir copy-import-dir
clean: clean-$(SRC_IMPORT_ROOT_DIR)
	if ( Test-Path $(SRC_ROOT_DIR) ) { ${RM} $(SRC_ROOT_DIR) }
clean-$(SRC_IMPORT_ROOT_DIR):
	if ( Test-Path $(SRC_IMPORT_ROOT_DIR) ) { ${RM} $(SRC_IMPORT_ROOT_DIR) }
create-src-root-dir:
	if ( -not (Test-Path $(SRC_ROOT_DIR)) ) { mkdir $(SRC_ROOT_DIR) }
copy-import-dir: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)

run:
	cscript $(VBAC_EXE) run /binary:$(abspath $(XLSM)) /entrypoint:$(ENTRYPOINT)

export: create-src-root-dir clean-$(SRC_IMPORT_ROOT_DIR)
	if (-not ( Test-Path $(SRC_ROOT_DIR)/$(XLSM_PARENT_DIR) )) { mkdir $(SRC_ROOT_DIR)/$(XLSM_PARENT_DIR) }
	cscript $(VBAC_EXE) decombine /source:$(SRC_ROOT_DIR)/$(XLSM_PARENT_DIR) /binary:$(XLSM_ABSPATH)
	Get-ChildItem -Recurse -Attributes !Directory $(MACROS_DIR)  | %{ nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) -Lu --overwrite $$_.FullName }

import: copy-import-dir
	Get-ChildItem -Recurse -Attributes !Directory $(SRC_IMPORT_ROOT_DIR)/$(XLSM_RELPATH)  | %{ nkf --ic=$(THIS_ENCODING) --oc=$(VBA_ENCODING) --overwrite $$_.FullName }
	cscript $(VBAC_EXE) combine /source:$(SRC_IMPORT_ROOT_DIR)/$(XLSM_PARENT_DIR) /binary:$(XLSM_ABSPATH)

unbind:
	cscript $(VBAC_EXE) clear /binary:$(XLSM_ABSPATH)


# macOS or linux
else

RM=rm -rf
TOUCH=touch
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
import: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
	find $(SRC_IMPORT_ROOT_DIR)/$(XLSM_RELPATH) -type f -print -exec nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) --overwrite {} \;
	./scripts/action_macos.scpt "import" $(HELPER_XLSM) $(XLSM_ABSPATH) $(abspath $(SRC_IMPORT_ROOT_DIR)/$(XLSM_RELPATH))

unbind:
	./scripts/action_macos.scpt "unbind" $(HELPER_XLSM) $(XLSM_ABSPATH)

export:
	$(RM) $(MACROS_DIR)
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$(XLSM_RELPATH) --dst_dir /code/$(MACROS_DIR)

clean:
	$(RM) $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
endif
