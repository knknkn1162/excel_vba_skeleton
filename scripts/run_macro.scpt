#!/usr/bin/env osascript -l JavaScript

function run(argv) {
    Application("Microsoft Excel.app").quit()
    path = argv[0];
    entrypoint = argv[1]
    try {
    	var xlApp = Application("Microsoft Excel.app");
    	xlApp.frontmost = true;
    	try {
    		var xlBook = xlApp.open(path);
            xlApp.runVBMacro(entrypoint);
            console.log("press button");
            var app = Application.currentApplication();
            app.includeStandardAdditions = true;
            app.displayAlert('finish!');
    	} finally {
    		if (xlBook != null) xlBook.close();
    	}
    } finally { 
    	if (xlApp != null) xlApp.quit();
    }
}
