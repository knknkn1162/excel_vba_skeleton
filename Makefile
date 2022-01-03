DIRS=excelvba9 \
	 excelvba1

XLSMS=$(filter-out $(wildcard */~*.xlsm), \
	  $(foreach dir, $(DIRS), $(wildcard $(dir)/*.xlsm)) \
)
TARGETS=$(basename $(XLSMS))
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
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor /code/$^ --dst_dir /code/src/$^
	git add $@
	-git commit -m "$(COMMIT_MSG) $@"

ifeq ("$(OS)", "Windows_NT")
ifeq (,$(wildcard ./src))
$(error "create ./src first!")
endif
import-%: %/
	cscript ./vbac/vbac.wsf combine /source:src/$^ /binary:$^
export-%: %/
	cscript ./vbac/vbac.wsf decombine /source:src/$^ /binary:$^
import: $(addprefix import-, $(DIRS))
export: $(addprefix export-, $(DIRS))
else
import:
	$(error "import command is not implemented")	
export:
	$(error "import command is not implemented")
endif
	

clean:
	$(RM) -r $(TARGETS)
