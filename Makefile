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

.PHONY: all imoprt export clean

all: export

ifeq ("$(OS)", "Windows_NT")
SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -Command
RM=rm -r -fo
ifeq (,$(wildcard $(SRC_ROOT_DIR)/))
DO_STUFF=create-src-root-dir
endif
.PHONY: create-src-root-dir copy-import-dir
clean-$(SRC_IMPORT_ROOT_DIR):
	if ( Test-Path $(SRC_IMPORT_ROOT_DIR) ) { ${RM} $(SRC_IMPORT_ROOT_DIR) }
clean: clean-$(SRC_IMPORT_ROOT_DIR)
	if ( Test-Path $(SRC_ROOT_DIR) ) { ${RM} $(SRC_ROOT_DIR) }
create-src-root-dir:
	mkdir $(SRC_ROOT_DIR)
copy-import-dir: clean-$(SRC_IMPORT_ROOT_DIR)
	cp -r $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)


import-%: % copy-import-dir
	# UTF-8 -> Shift_JIS and combine
	Get-ChildItem -Recurse -Attributes !Directory $(SRC_IMPORT_ROOT_DIR)/$< | %{ nkf --ic=$(THIS_ENCODING) --oc=$(VBA_ENCODING) --overwrite $$_.FullName }
	cscript ./vbac/vbac.wsf combine /source:$(SRC_IMPORT_ROOT_DIR)/$< /binary:$<

export-%: % $(DO_STUFF) clean-$(SRC_IMPORT_ROOT_DIR)
	cscript ./vbac/vbac.wsf decombine /source:$(SRC_ROOT_DIR)/$< /binary:$<
	# Shift_JIS -> UTF-8, CRLF -> LU
	Get-ChildItem -Recurse -Attributes !Directory $(SRC_ROOT_DIR)/$< | %{ nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) -Lu --overwrite $$_.FullName }

import: $(addprefix import-, $(DIRS))
export: $(addprefix export-, $(DIRS))

unbind-%: %
	cscript ./vbac/vbac.wsf clear /binary:$<
unbind: $(addprefix unbind-, $(DIRS))

else
RM=rm -rf
import:
	$(error "import command is not implemented")	
export: $(TARGETS)
%: %.xlsm
	$(RM) $@
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$^ --dst_dir /code/$(SRC_ROOT_DIR)/$^
clean:
	$(RM) $(SRC_ROOT_DIR) $(SRC_IMPORT_ROOT_DIR)
endif
