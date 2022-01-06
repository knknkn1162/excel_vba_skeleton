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

# define macro
ifeq ("$(OS)", "Windows_NT")

define define-run-commands
run-$(1)-%: $(1)/%
	cscript $(VBAC_EXE) run /target:$$^
run-$(1): $(1) $(addprefix run-$(1)-, $(notdir $(wildcard $(1)/*.xlsm)))
endef

else
endif

# common commands

.PHONY: all imoprt export clean
all: export
## run commands
$(foreach dir, $(DIRS), \
$(eval $(call define-run-commands,$(dir))) \
)

# OS dep. commands
ifeq ("$(OS)", "Windows_NT")
SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -Command
RM=rm -r -fo
ifeq (,$(wildcard $(SRC_ROOT_DIR)/))
DO_STUFF=create-src-root-dir
endif
.PHONY: create-src-root-dir copy-import-dir
clean: clean-$(SRC_IMPORT_ROOT_DIR)
	if ( Test-Path $(SRC_ROOT_DIR) ) { ${RM} $(SRC_ROOT_DIR) }
clean-$(SRC_IMPORT_ROOT_DIR):
	if ( Test-Path $(SRC_IMPORT_ROOT_DIR) ) { ${RM} $(SRC_IMPORT_ROOT_DIR) }
create-src-root-dir:
	mkdir $(SRC_ROOT_DIR)
copy-import-dir: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)


import-all: $(addprefix import-, $(DIRS))
import-%: % copy-import-dir
	# UTF-8 -> Shift_JIS and combine
	Get-ChildItem -Recurse -Attributes !Directory $(SRC_IMPORT_ROOT_DIR)/$< | %{ nkf --ic=$(THIS_ENCODING) --oc=$(VBA_ENCODING) --overwrite $$_.FullName }
	cscript $(VBAC_EXE) combine /source:$(SRC_IMPORT_ROOT_DIR)/$< /binary:$<


export-all: $(addprefix export-, $(DIRS))
export-%: % $(DO_STUFF) clean-$(SRC_IMPORT_ROOT_DIR)
	cscript $(VBAC_EXE) decombine /source:$(SRC_ROOT_DIR)/$< /binary:$<
	# Shift_JIS -> UTF-8, CRLF -> LU
	Get-ChildItem -Recurse -Attributes !Directory $(SRC_ROOT_DIR)/$< | %{ nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) -Lu --overwrite $$_.FullName }

unbind-%: %
	cscript $(VBAC_EXE) clear /binary:$<
unbind-all: $(addprefix unbind-, $(DIRS))

# macOS or linux
else

RM=rm -rf
# Mac OS only
ifeq ("$(shell uname)", "Darwin")
run:
	./scripts/run_macro.scpt $(abspath $(XLSM))
else
run:
	$(error "run command is not implemented")
endif
import:
	$(error "import command is not implemented")

export: $(TARGETS)
%: %.xlsm
	$(RM) $@
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$^ --dst_dir /code/$(SRC_ROOT_DIR)/$^
clean:
	$(RM) $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
endif
