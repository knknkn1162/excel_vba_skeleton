#!/usr/bin/env osascript -l JavaScript

function run(argv) {
    Application("Microsoft Excel.app").quit();
    action = argv[0]
    originPath = argv[1];
    dst = argv[2];
    src_dir = argv[3];
    try {
    	var xlApp = Application("Microsoft Excel.app");
    	xlApp.frontmost = true;
    	try {
    		var xlBook = xlApp.open(originPath);
    		xlApp.runVBMacro(action, {arg1: dst, arg2: src_dir});
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
