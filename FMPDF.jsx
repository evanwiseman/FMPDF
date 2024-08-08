// Get the active document.
var doc = app.ActiveDoc;

function getFileNameFromPath(filePath) {
    // Split the file path by the directory separator
    var parts = filePath.split(/[\\\/]/);
    // Extract the last part of the array, which represents the file name
    var fileName = parts[parts.length - 1];
    return fileName;
}

function importPDFPage(path, pageNum) {
    if((doc.ObjectValid() == true) && (path != "") && (pageNum >= 1))
    {
        // Get users selected text location
        var textLoc = new TextLoc();
        textLoc = doc.TextSelection.beg;

        doc.AddText(textLoc, "\n");

        // Create import parameters
        var importParams = GetImportDefaultParams();
        var index = GetPropIndex(importParams, Constants.FS_HowToImport);
        if (index > -1) importParams[index].propVal.ival = Constants.FV_DoByRef;
        index = GetPropIndex(importParams, Constants.FS_FitGraphicInSelectedRect);
        if(index > -1) importParams[index].propVal.ival = false;
        index = GetPropIndex(importParams, Constants.FS_GraphicDpi);
        if(index > -1) importParams[index].propVal.ival = 600;
        index = GetPropIndex(importParams, Constants.FS_PDFPageNum);
        if(index > -1) importParams[index].propVal.ival = pageNum;

        // Create import return parameters
        var importReturnParams = new PropVals();

        // Import from path at selected text location
        doc.Import(textLoc, path, importParams, importReturnParams);

        var frame = null;
        var graphic = null;
        graphic = doc.FirstSelectedGraphicInDoc;
        if (graphic.ObjectValid() && graphic.constructor.name == "Inset") 
        {
            frame = graphic.FrameParent;
            
            if (!frame.ObjectValid() || frame.constructor.name != "AFrame")
            {
                frame = doc.NewAnchoredAFrame();
                graphic.FrameParent = frame;
            }

            graphic.LocX = 65535;
            graphic.LocY = 65535;
            graphic.Height *= 0.952;
            graphic.Width *= 0.952;

            frame.Height = (graphic.Height + (2 * 65535));
            frame.Width = (graphic.Width + (2 * 65535));
            frame.Alignment = Constants.FV_ALIGN_CENTER;
            frame.BorderWidth = 1;
        }
        else if (graphic.ObjectValid() && graphic.constructor.name == "AFrame") 
        {
            frame = graphic;
            graphic.LocX = 65535;
            graphic.LocY = 65535;
            graphic.FirstGraphicInFrame.Height *= 0.952;
            graphic.FirstGraphicInFrame.Width *= 0.952;

            frame.Height = (graphic.FirstGraphicInFrame.Height + (2 * 65535));
            frame.Width = (graphic.FirstGraphicInFrame.Width + (2 * 65535));
            frame.Alignment = Constants.FV_ALIGN_CENTER;
            frame.BorderWidth = 1;
        }
        
        if (frame == null)
        {
            alert("Error: Couldn't select frame.");
            return;
        }

        doc.AddText(textLoc, "\n");
        var textRange = new TextRange();
        textRange.beg.obj = textRange.end.obj = textLoc.obj.NextPgfInFlow;
        textRange.beg.offset = 0;
        textRange.end.offset = Constants.FV_OBJ_END_OFFSET;
        doc.TextSelection = textRange;
    }
}

function countPagesInPDF(filepath) {
    try {
        // Create a file object
        var file = new File(filepath);

        // Open the file
        if (!file.open("r")) {
            alert("Error: Failed to open the file");
            return -1;
        }

        // Initialize variables
        var pageCount = 0;
        var regex = /\/Type\s*\/Page\b/g;

        // Read the file line by line
        var line;
        while (!file.eof) {
            line = file.readln();
            if (regex.test(line)) {
                pageCount++;
            }
        }

        // Close the file
        file.close();

        return pageCount;
    } catch (error) {
        // Handle any errors
        alert("Error: " + error.message);
        return -1; // Return -1 to indicate an error
    }
}

function getFilePathsFromFolder(folderPath) {
    try {
        var folder = new Folder(folderPath);
        var fileList = [];

        if (!folder.exists) {
            alert("Error: Folder does not exist");
            return []
        }

        var files = folder.getFiles();
        for (var i = 0; i < files.length; i++) {
            if (files[i] instanceof File && files[i].fullName.match(/\.pdf$/i)) {
                fileList.push(files[i].fsName);
            }
        }

        return fileList;
    } catch (error) {
        // Handle any errors
        alert("An error occurred: " + error.message);
        return []; // Return an empty array to indicate an error
    }
}

function main() {
    if (doc.ObjectValid() == true) {
        var i = 0;

        var dlg = new Window("dialog", "Select Option");
        dlg.alignChildren = "center";

        var btnFile = dlg.add("button", undefined, "File");
        btnFile.onClick = function() {
            dlg.close();
            try {
                var filepath = ChooseFile("Browse for graphic file", doc.Name, "", Constants.FV_ChooseSelect);
                selectFile(filepath);
            } catch(e) {
                alert("Error: Failed selecting file.");
            }
        };

        var btnFolder = dlg.add("button", undefined, "Folder");
        btnFolder.onClick = function() {
            dlg.close();
            try {
                var folderpath = ChooseFile("Browse for folder", doc.Name, "", Constants.FV_ChooseOpenDir);
                selectFolder(folderpath);
            } catch(e) {
                alert("Error: Failed selecting directory.");
            }
        };

        dlg.show();

        function selectFile(filepath) {
            try {
                if (File(filepath).exists) {

                    i = countPagesInPDF(filepath);
                    while(i >= 1) {
                        importPDFPage(filepath, i);
                        $.writeln("Page: " + i);
                        i = i - 1;
                    }
                    textLoc = doc.TextSelection.beg;

                    if (textLoc.obj.NextPgfInFlow !== undefined && textLoc.obj.NextPgfInFlow !== null) {
                        // Move text selection to next paragrap
                        name = getFileNameFromPath(filepath);
                        doc.AddText(textLoc, name + "\n");
                    } else {
                        alert("Error: Bad insertion point.")
                        return;
                    }
                    
                }
            } catch(e) {
                alert("Error: " + e);
            }
        }

        function selectFolder(folderpath) {
            $.writeln(folderpath);
            var files = getFilePathsFromFolder(folderpath);
            for (var i = files.length - 1; i >= 0; i--) {
                $.writeln(files[i]);
                selectFile(files[i]);
            }
        }   
    }
}

main();