// Enums
function Enum(constantsList) {
    for (var i in constantsList) {
        this[constantsList[i]] = i;
    }
}

// Delete a prior non-NULL value
// Change a non-NULL value to a different non-NULL value
// Change a NULL value to a non-NULL value
var EditType = new Enum(['DELETE', 'CHANGE', 'INIT']);

// Generic "on edit" by user
// NOTE: Minimize the code here as it will affect overall Sheet performance
function onEdit(e) {
    'use strict';
    var s = SpreadsheetApp.getActiveSheet();

    // Edits happen on a "grid" of 'x by y' cells
    // So, we will iterate over the changes, one by one
    var active_cells = s.getActiveRange();

    var start_row = active_cells.getRow();
    var start_column = active_cells.getColumn();

    var width = active_cells.getWidth();
    var height = active_cells.getHeight();

    var i;
    var j;

    var edit_type = EditType.DELETE;
    if (width == 1 && height == 1) {
        if ((e.oldValue == "" || e.oldValue == undefined) && active_cells.getValue() !== "") {
            edit_type = EditType.INIT;
        }
        else if ((e.oldValue != "" && e.oldValue != undefined) && active_cells.getValue() === "") {
            edit_type = EditType.DELETE;
        }
        else {
            edit_type = EditType.CHANGE;
        }
    }
    else {
        if (active_cells.getValue() === "") {
            edit_type = EditType.DELETE;
        }
        else {
            edit_type = EditType.CHANGE;
        }
    }

    // Go through each change
    for (i = 0; i < width; i++) {
        for (j = 0; j < height; j++) {
            var row = start_row + j;
            var column = start_column + i;

            var current_cell = s.getRange(row, column);

            if (row == 1) // Localization languages
            {
                if (column > 1) {
                    // Top bar list of languages changing
                    changeLanguageList(s, current_cell, edit_type);
                }
            }
            else if (column == 1) {
                // StringId changing
                changeStringId(s, current_cell, edit_type);
            }
            else if (column == 2) {
                // Primary language text changing
                changeSourceText(s, current_cell, edit_type);
            }
            else {
                // Likely translated text changing
                changeTranslatedText(s, current_cell, edit_type);
            }
        }
    }
}

function informDuplicate(x, y) {
    var ui = SpreadsheetApp.getUi(); // Same variations.

    ui.alert(
    'Warning',
    'Row ' + x + ' and row ' + y + ' both have the same id',
    ui.ButtonSet.OK);
}

// s - SpreadsheetApp
// r - Cell being changed
// EditType.INIT, EditType.DELETE, EditType.CHANGE  
function changeStringId(s, r, edit_type) {
    var value = r.getValue();

    if (value !== "" && (edit_type == EditType.CHANGE || edit_type == EditType.INIT)) {
        var current_row = r.getRow();

        // Determine size of the sheet
        var start_row = 2;
        var data_range = s.getDataRange();
        var end_row = data_range.getLastRow();

        var values = s.getRange(start_row, 1, end_row - start_row + 1, 1).getValues();
        var new_row = current_row - start_row;

        var row = 0;
        while (row <= end_row - start_row) {
            if (row !== new_row) {
                if (value === values[row][0]) {
                    informDuplicate(current_row, row + start_row);
                }
            }
            row += 1;
        }
    }
}

// s - SpreadsheetApp
function getLanguageListCount(s) {
    // Location of the primary language
    var primary_language = s.getRange(1, 1);
    var count = 0;

    var nextCell = primary_language.offset(0, 1);
    while (nextCell.getValue() !== "") {
        count += 1;
        nextCell = nextCell.offset(0, 1);
    }

    return count;
}

function askInvalid() {
    var ui = SpreadsheetApp.getUi(); // Same variations.

    var result = ui.alert(
     'Please confirm',
     'Does the updated text require new translation?',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
        // User clicked "Yes".
        return true;
    }
    else {
        // User clicked "No" or X in the title bar.
        return false;
    }
}

// s - SpreadsheetApp
// r - Cell being changed
// EditType.INIT, EditType.DELETE, EditType.CHANGE  
function changeSourceText(s, r, edit_type) {
    //var s = SpreadsheetApp.getActiveSheet();
    var row = r.getRow();

    // Don't process this change if there is no string id
    var string_id = s.getRange(row, 1).getValue();
    if (string_id === "") {
        return;
    }

    var lang_count = getLanguageListCount(s);
    //SpeadsheetApp.getUi().alert('lang', lang_count, SpeadsheetApp.ButtonSet.YES_NO);

    var invalidate = true;
    if (edit_type == EditType.CHANGE) {
        invalidate = askInvalid();
    }

    if (invalidate) {
        lang_count--; // Skip the primary language (English)
        var nextCell = r.offset(0, 1);

        // Constant
        var source_language = s.getRange(1, 2).getValue();
        var source_text = s.getRange(row, 2).getValue();

        while (lang_count > 0) {
            // If it is red or doesn't have a value set, then update the Google Translate
            if (nextCell.getBackground() == "#ff0000" || nextCell.getValue() === "") {
                var column = nextCell.getColumn();
                var dest_language = s.getRange(1, column).getValue();

                var translatedText = LanguageApp.translate(source_text, source_language, dest_language);
                nextCell.setValue(translatedText);
                nextCell.setBackground("red");
            }
            else {
                nextCell.setBackground("orange");
            }

            nextCell = nextCell.offset(0, 1);
            lang_count -= 1;
        };
    }
}

// s - SpreadsheetApp
// r - Cell being changed
// EditType.INIT, EditType.DELETE, EditType.CHANGE  
function changeTranslatedText(s, r, edit_type) {
    var lang_count = getLanguageListCount(s);
    var column = r.getColumn();
    var row = r.getRow();
    var string_id = s.getRange(row, 1).getValue();

    if (string_id !== "" && column > 1 && column <= lang_count + 1) // Translated languages
    {
        if (r.getValue() !== "") {
            r.setBackground("green");
        }
        else {
            r.setBackground("orange");
        }
    }
}

function askGoogleTranslate() {
    var ui = SpreadsheetApp.getUi(); // Same variations.

    var result = ui.alert(
     'New Language Added',
     'Do you want to add Google Translated text to the language?',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
        // User clicked "Yes".
        return true;
    }
    else {
        // User clicked "No" or X in the title bar.
        return false;
    }
}

// s - SpreadsheetApp
// r - Cell being changed
// EditType.INIT, EditType.DELETE, EditType.CHANGE  
function changeLanguageList(s, r, edit_type) {
    if (edit_type == EditType.INIT) {
        if (askGoogleTranslate()) {
            var data_range = s.getDataRange();
            var start_row = 2;
            var end_row = data_range.getLastRow();

            var row = start_row;
            var column = r.getColumn();

            var source_language = s.getRange(1, 2).getValue();
            var dest_language = s.getRange(1, column).getValue();

            var id = s.getRange(row, 1, end_row - start_row + 1, 1).getValues();
            var text = new Array(end_row - start_row + 1, 1);  // must be a 2d array
            text = s.getRange(row, 2, end_row - start_row + 1, 1).getValues();

            var loc_text = s.getRange(start_row, column, end_row - start_row + 1, 1);
            loc_text.setBackground('red');

            // Bake out the values
            // NOTE: Cannot iterate through more than 500 cells as GAS has a limit of 500 calls to LanguageApp.translate()
            if (end_row - start_row + 1 <= 500) {
                while (row <= end_row) {
                    if (id[row - start_row] !== "" && text[row - start_row][0] !== "") {
                        // Translate only have single element calls
                        text[row - start_row][0] = LanguageApp.translate(text[row - start_row][0], source_language, dest_language);
                    }
                    row += 1;
                }

                // Do a single big set call
                loc_text.setValues(text);
            }
            else {
                // I prefer the backed text solution above yet this works
                loc_text.setFormula('=GOOGLETRANSLATE(B2, "' + source_language + '", "' + dest_language + '")');
            }
        }
    }
}

// s - SpreadsheetApp
function getLanguageList() {
    var s = SpreadsheetApp.getActiveSheet();

    // Location of the primary language
    var primary_language = s.getRange(1, 1);
    var name_array = [];

    var nextCell = primary_language.offset(0, 1);
    while (nextCell.getValue() !== "") {
        name_array.push(nextCell.getValue());
        nextCell = nextCell.offset(0, 1);
    }

    return name_array;
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('VLGTrans')
      .addItem('Export', 'ExportAllLanguages')
      .addToUi();
}

function ExportAllLanguages() {
    var s = SpreadsheetApp.getActiveSheet();
    var languages = getLanguageList();
    var language_count = languages.length;

    // Determine size of the sheet
    var data_range = s.getDataRange();
    var start_row = 2;
    var end_row = data_range.getLastRow();

    var cur_lang = 0;

    for (cur_lang = 0; cur_lang < language_count; cur_lang++) {
        var text_file = "{\n";
        var row = start_row;
        var column = 2 + cur_lang;

        var ids = s.getRange(start_row, 1, end_row - start_row + 1, 1).getValues();
        var lang_strings = s.getRange(start_row, column, end_row - start_row + 1, 1).getValues();

        while (row <= end_row) {
            if (ids[row - start_row] !== "") {
                // Allow both Alt+Enter and \n in strings
                var res1 = lang_strings[row - start_row].toString().replace(/(?:\r\n|\r|\n)/g, '\\n'); // replace real Alt-Enter with a \n

                // Allow both \" and " without any problems
                var res2 = res1.replace(/\\\"/g, '\"'); // replace \" with a " (so, all are just a ")
                var text = res2.replace(/\"/g, '\\\"'); // make all single " a \"

                text_file += '"' + ids[row - start_row] + '":"' + text + '"';
                if (row < end_row) {
                    text_file += ",";
                }
                text_file += "\n";
            }

            row += 1;
        }
        text_file += "}";

        DriveApp.createFile(s.getName() + '_' + languages[cur_lang], text_file, MimeType.PLAIN_TEXT);
    }
}


