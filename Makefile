DIRS=excelvba9 \
	 excelvba1

XLSMS=$(filter-out $(wildcard */~*.xlsm), \
	  $(foreach dir, $(DIRS), $(wildcard $(dir)/*.xlsm)) \
)
TARGETS=$(basename $(XLSMS))
SRC_ROOT_DIR=./src
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
create-src-root-dir:
	mkdir ${SRC_ROOT_DIR}
ifeq (,$(wildcard ${SRC_ROOT_DIR}/))
DO_STUFF=create-src-root-dir
endif

import-%: %/
	cscript ./vbac/vbac.wsf combine /source:${SRC_ROOT_DIR}/$^ /binary:$^
export-%: $(DO_STUFF) %/
	cscript ./vbac/vbac.wsf decombine /source:${SRC_ROOT_DIR}/$^ /binary:$^
import: $(addprefix import-, $(DIRS))
export: $(addprefix export-, $(DIRS))
else
import:
	$(error "import command is not implemented")	
export:
	$(error "import command is not implemented")
endif
	

clean:
	$(RM) -r ${SRC_ROOT_DIR}
