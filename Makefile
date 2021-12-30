XLSMS=$(filter-out $(wildcard */~*.xlsm), $(wildcard excelvba9/*.xlsm))
TARGETS=$(basename $(XLSMS))

test:
	@echo $(TARGETS)
all: $(TARGETS)

%: %.xlsm
	@echo "build $^"
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor -- /code/$^
	git add $@
	-git commit -m "impl $@"
