XLSMS=$(filter-out $(wildcard */~*.xlsm), $(wildcard excelvba9/*.xlsm))
TARGETS=$(basename $(XLSMS))

.PHONY: all push commit clean

all: commit push

push:
	git push

commit: $(TARGETS)

print-%: ; @echo $* = $($*)

%: %.xlsm
	@echo "build $^"
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor -- /code/$^
	git add $@
	-git commit -m "impl $@"

clean:
	find ./excelvba9 -maxdepth 1 -mindepth 1 -type d -exec rm -rf '{}' \;
