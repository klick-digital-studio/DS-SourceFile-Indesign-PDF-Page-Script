// Place PDF with Sized Pages - Global Version
// Fixed for localization issues (Portuguese/International units)

#target indesign

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }

    var inddDoc = app.activeDocument;
    
    // 1. Save original user settings
    var oldUserInteraction = app.scriptPreferences.userInteractionLevel;
    var oldXUnits = inddDoc.viewPreferences.horizontalMeasurementUnits;
    var oldYUnits = inddDoc.viewPreferences.verticalMeasurementUnits;
    var oldOrigin = inddDoc.viewPreferences.rulerOrigin;

    // 2. Set script-safe environment
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    
    // Force Points and Page Origin to ensure math is consistent regardless of language
    inddDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    try {
        var pdfFile = File.openDialog("Select the PDF file to place", "*.pdf", false);
        if (!pdfFile) return;

        var pageCounter = 1;
        var firstPlacedPDFPageNum = -1; 

        while (true) {
            var targetPage;

            // Check if we can use the first page or need a new one
            if (pageCounter === 1 && inddDoc.pages[0].allPageItems.length === 0 && inddDoc.pages.length === 1) {
                targetPage = inddDoc.pages[0];
            } else {
                targetPage = inddDoc.pages.add(LocationOptions.AT_END);
            }

            // PDF Import Preferences
            app.pdfPlacePreferences.pageNumber = pageCounter;
            app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_MEDIA;

            // Place PDF
            var placedItemArray = targetPage.place(pdfFile, [0, 0]);

            if (placedItemArray.length === 0) {
                targetPage.remove();
                break;
            }

            var placedItem = placedItemArray[0];
            var currentPDFPageNum = placedItem.pdfAttributes.pageNumber;

            // Stop logic: If InDesign loops back to page 1 of the PDF
            if (pageCounter === 1) {
                firstPlacedPDFPageNum = currentPDFPageNum;
            } else if (currentPDFPageNum === firstPlacedPDFPageNum) {
                targetPage.remove(); 
                break;
            }

            // Calculate dimensions based on Geometric Bounds
            // InDesign returns [y1, x1, y2, x2]
            var pdfFrame = placedItem.parent;
            var pdfBounds = pdfFrame.geometricBounds; 

            var pdfWidth = pdfBounds[3] - pdfBounds[1];
            var pdfHeight = pdfBounds[2] - pdfBounds[0];
            
            // Reframe the InDesign page to match PDF dimensions
            // Reframe uses [ [x1, y1], [x2, y2] ]
            var newPageBounds = [[0, 0], [pdfWidth, pdfHeight]];
            targetPage.reframe(CoordinateSpaces.PAGE_COORDINATES, newPageBounds);

            // Ensure the PDF frame is perfectly aligned at top-left
            pdfFrame.move([0, 0]);

            pageCounter++;
        }

        alert("Success! Placed " + (pageCounter - 1) + " pages.");

    } catch (e) {
        alert("An error occurred: " + e.message + " (Line: " + e.line + ")");
    } finally {
        // 3. ALWAYS restore the user's original settings
        inddDoc.viewPreferences.horizontalMeasurementUnits = oldXUnits;
        inddDoc.viewPreferences.verticalMeasurementUnits = oldYUnits;
        inddDoc.viewPreferences.rulerOrigin = oldOrigin;
        app.scriptPreferences.userInteractionLevel = oldUserInteraction;
    }
}

main();// Place PDF with Sized Pages - Global Version
// Fixed for localization issues (Portuguese/International units)

#target indesign

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }

    var inddDoc = app.activeDocument;
    
    // 1. Save original user settings
    var oldUserInteraction = app.scriptPreferences.userInteractionLevel;
    var oldXUnits = inddDoc.viewPreferences.horizontalMeasurementUnits;
    var oldYUnits = inddDoc.viewPreferences.verticalMeasurementUnits;
    var oldOrigin = inddDoc.viewPreferences.rulerOrigin;

    // 2. Set script-safe environment
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    
    // Force Points and Page Origin to ensure math is consistent regardless of language
    inddDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    try {
        var pdfFile = File.openDialog("Select the PDF file to place", "*.pdf", false);
        if (!pdfFile) return;

        var pageCounter = 1;
        var firstPlacedPDFPageNum = -1; 

        while (true) {
            var targetPage;

            // Check if we can use the first page or need a new one
            if (pageCounter === 1 && inddDoc.pages[0].allPageItems.length === 0 && inddDoc.pages.length === 1) {
                targetPage = inddDoc.pages[0];
            } else {
                targetPage = inddDoc.pages.add(LocationOptions.AT_END);
            }

            // PDF Import Preferences
            app.pdfPlacePreferences.pageNumber = pageCounter;
            app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_MEDIA;

            // Place PDF
            var placedItemArray = targetPage.place(pdfFile, [0, 0]);

            if (placedItemArray.length === 0) {
                targetPage.remove();
                break;
            }

            var placedItem = placedItemArray[0];
            var currentPDFPageNum = placedItem.pdfAttributes.pageNumber;

            // Stop logic: If InDesign loops back to page 1 of the PDF
            if (pageCounter === 1) {
                firstPlacedPDFPageNum = currentPDFPageNum;
            } else if (currentPDFPageNum === firstPlacedPDFPageNum) {
                targetPage.remove(); 
                break;
            }

            // Calculate dimensions based on Geometric Bounds
            // InDesign returns [y1, x1, y2, x2]
            var pdfFrame = placedItem.parent;
            var pdfBounds = pdfFrame.geometricBounds; 

            var pdfWidth = pdfBounds[3] - pdfBounds[1];
            var pdfHeight = pdfBounds[2] - pdfBounds[0];
            
            // Reframe the InDesign page to match PDF dimensions
            // Reframe uses [ [x1, y1], [x2, y2] ]
            var newPageBounds = [[0, 0], [pdfWidth, pdfHeight]];
            targetPage.reframe(CoordinateSpaces.PAGE_COORDINATES, newPageBounds);

            // Ensure the PDF frame is perfectly aligned at top-left
            pdfFrame.move([0, 0]);

            pageCounter++;
        }

        alert("Success! Placed " + (pageCounter - 1) + " pages.");

    } catch (e) {
        alert("An error occurred: " + e.message + " (Line: " + e.line + ")");
    } finally {
        // 3. ALWAYS restore the user's original settings
        inddDoc.viewPreferences.horizontalMeasurementUnits = oldXUnits;
        inddDoc.viewPreferences.verticalMeasurementUnits = oldYUnits;
        inddDoc.viewPreferences.rulerOrigin = oldOrigin;
        app.scriptPreferences.userInteractionLevel = oldUserInteraction;
    }
}

main();