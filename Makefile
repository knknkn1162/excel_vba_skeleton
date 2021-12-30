DIRS=excelvba9 \
	 excelvba1

XLSMS=$(filter-out $(wildcard */~*.xlsm), \
	  $(foreach dir, $(DIRS), $(wildcard $(dir)/*.xlsm)) \
)
TARGETS=$(basename $(XLSMS))
COMMIT_MSG=implement

.PHONY: all push commit clean

all: commit push

push:
	git push

commit: $(TARGETS)

print-%:
	@echo $* = $($*)

%: %.xlsm
	$(RM) -r $@
	@echo "build $^"
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor -- /code/$^
	git add $@
	-git commit -m "$(COMMIT_MSG) $@"

clean:
	$(RM) -r $(TARGETS)
