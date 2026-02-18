// Place PDF with Sized Pages - V4 (Global Fixed)
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
    inddDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    inddDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    try {
        var pdfFile = File.openDialog("Select the PDF file to place", "*.pdf", false);
        if (!pdfFile) return;

        var pageCounter = 1;
        var firstPlacedPDFPageNum = -1; 
        var finished = false;

        while (!finished) {
            var targetPage;

            // Use existing first page if it's empty, otherwise add new
            if (pageCounter === 1 && inddDoc.pages[0].allPageItems.length === 0 && inddDoc.pages.length === 1) {
                targetPage = inddDoc.pages[0];
            } else {
                targetPage = inddDoc.pages.add(LocationOptions.AT_END);
            }

            // Set PDF Import Preferences
            app.pdfPlacePreferences.pageNumber = pageCounter;
            app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_MEDIA;

            // Place PDF on the page
            var placedItemArray = targetPage.place(pdfFile, [0, 0]);
            
            if (placedItemArray.length === 0) {
                targetPage.remove();
                finished = true;
                break;
            }

            var placedItem = placedItemArray[0];
            var currentPDFPageNum = placedItem.pdfAttributes.pageNumber;

            // Logic to detect if we have looped back to page 1 (meaning the PDF is finished)
            if (pageCounter === 1) {
                firstPlacedPDFPageNum = currentPDFPageNum;
            } else if (currentPDFPageNum === firstPlacedPDFPageNum) {
                targetPage.remove(); 
                finished = true;
                break;
            }

            // --- REFRAME PAGE TO MATCH PDF ---
            var pdfFrame = placedItem.parent;
            var pdfBounds = pdfFrame.geometricBounds; // [y1, x1, y2, x2]

            var pdfWidth = pdfBounds[3] - pdfBounds[1];
            var pdfHeight = pdfBounds[2] - pdfBounds[0];
            
            // Apply new size to the page
            var newPageBounds = [[0, 0], [pdfWidth, pdfHeight]];
            targetPage.reframe(CoordinateSpaces.PAGE_COORDINATES, newPageBounds);

            // Center the PDF precisely
            pdfFrame.move([0, 0]);

            pageCounter++;
        }

        // Only alert after the loop is fully broken
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
        alert("Finished! Placed " + (pageCounter - 1) + " pages.");

    } catch (e) {
        alert("An error occurred: " + e.message + " (Line: " + e.line + ")");
    } finally {
        // 3. ALWAYS restore settings
        inddDoc.viewPreferences.horizontalMeasurementUnits = oldXUnits;
        inddDoc.viewPreferences.verticalMeasurementUnits = oldYUnits;
        inddDoc.viewPreferences.rulerOrigin = oldOrigin;
        app.scriptPreferences.userInteractionLevel = oldUserInteraction;
    }
}

main();