%: %.xlsm
	docker run -it -v $(PWD):/code --rm knknkn1162/vba_extractor -- /code/$^
	git add $@
	git commit -m "impl $@"
	git push
