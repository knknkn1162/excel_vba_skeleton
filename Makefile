DIRS=excelvba9 \
	 excelvba1

XLSMS=$(filter-out $(wildcard */~*.xlsm), \
	  $(foreach dir, $(DIRS), $(wildcard $(dir)/*.xlsm)) \
)
TARGETS=$(basename $(XLSMS))
SRC_ROOT_DIR=src
VBA_ENCODING=Shift_JIS
THIS_ENCODING=UTF-8
SRC_IMPORT_ROOT_DIR=${SRC_ROOT_DIR}_${VBA_ENCODING}
COMMIT_MSG=implement

.PHONY: all push commit imoprt export clean

all: commit push

push:
	git push

commit: $(TARGETS)

print-%:
	@echo $* = $($*)

%: %.xlsm
	$(RM) -r $@
	@echo "build $^"
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$^ --dst_dir /code/${SRC_ROOT_DIR}/$^
	git add ${SRC_ROOT_DIR}/$^
	-git commit -m "$(COMMIT_MSG) $@"

ifeq ("$(OS)", "Windows_NT")
SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -Command
create-src-root-dir:
	mkdir ${SRC_ROOT_DIR}
copy-import-dir:
	cp -r ${SRC_ROOT_DIR} ${SRC_IMPORT_ROOT_DIR}
ifeq (,$(wildcard ${SRC_ROOT_DIR}/))
DO_STUFF=create-src-root-dir
endif

# UTF-8 -> Shift_JIS and combine
import-%: %
	nkf --ic=$(THIS_ENCODING) --oc=$(VBA_ENCODING) --overwrite ${SRC_IMPORT_ROOT_DIR}/$</*/*
	cscript ./vbac/vbac.wsf combine /source:${SRC_IMPORT_ROOT_DIR}/$^ /binary:$^

# Shift_JIS -> UTF-8, CRLF -> LU
export-%: % $(DO_STUFF)
	cscript ./vbac/vbac.wsf decombine /source:${SRC_ROOT_DIR}/$< /binary:$<
	nkf --ic=$(VBA_ENCODING) --oc=$(THIS_ENCODING) -Lu --overwrite ${SRC_ROOT_DIR}/$</*/*
import: copy-import-dir $(addprefix import-, $(DIRS))
export: $(addprefix export-, $(DIRS))
	rm -r ${SRC_IMPORT_ROOT_DIR}
clean:
	rm -r ${SRC_ROOT_DIR}
	rm -r ${SRC_IMPORT_ROOT_DIR}
else
import:
	$(error "import command is not implemented")	
export:
	$(error "import command is not implemented")
clean:
	$(RM) -r ${SRC_ROOT_DIR}
endif
