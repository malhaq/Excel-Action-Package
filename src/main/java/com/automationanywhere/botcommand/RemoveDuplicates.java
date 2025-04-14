package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "RemoveDuplicates", label = "[[RemoveDuplicates.label]]", node_label = "[[RemoveDuplicates.node_label]]", description = "[[RemoveDuplicates.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[RemoveDuplicates.return_label]]", return_type = STRING, return_required = true)

public class RemoveDuplicates {
    private static final Logger LOGGER = Logger.getLogger(RemoveDuplicates.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input Excel File Path",description = "Path to the input Excel file.")
            @NotEmpty
            String inputFilePath,


            @Idx(index = "2", type = TEXT)
            @Pkg(label = "Output Excel File Path",description = "Path where the modified Excel file will be saved.")
            @NotEmpty
            String outputFilePath,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Specified Column (e.g., A)",description = "The column letter used for duplicate detection.")
            @NotEmpty
            String column


    ) {


        Set<String> seenValues = new HashSet<>();
        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {
            Sheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = sheet.getLastRowNum();
            int startColumn = ExcelUtils.columnLetterToIndex(column);

            for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                Cell cell = row.getCell(startColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String val;
                if (cell.getCellType() == CellType.NUMERIC) {
                    val = String.valueOf(cell.getNumericCellValue());
                } else {
                    val = cell.getStringCellValue();
                }
                if (!seenValues.add(val)) {
                    LOGGER.log(Level.INFO, "Duplication found on row {0}", rowNum + 1);
                    sheet.removeRow(row);
                    if (rowNum < endRow) {
                        sheet.shiftRows(rowNum + 1, endRow, -1);
                        endRow--;
                        rowNum--;
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            return new StringValue("Duplicate detection completed all duplicates have been deleted. Output Saved as: " + outputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }




}

