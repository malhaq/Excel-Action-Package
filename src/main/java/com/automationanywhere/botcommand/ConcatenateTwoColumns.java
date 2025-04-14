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

import java.io.*;
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ConcatenateTwoColumns", label = "[[ConcatenateTwoColumns.label]]", node_label = "[[ConcatenateTwoColumns.node_label]]", description = "[[ConcatenateTwoColumns.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[ConcatenateTwoColumns.return_label]]", return_type = STRING, return_required = true)

public class ConcatenateTwoColumns {
    private static final Logger LOGGER = Logger.getLogger(ConcatenateTwoColumns.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input Excel File Path", description = "Path to the input Excel file.")
            @NotEmpty
            String inputFilePath,

//            @Idx(index = "2", type = TEXT)
//            @Pkg(label = "Output Excel File Path", description = "Path where the modified Excel file will be saved.")
//            @NotEmpty
//            String outputFilePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "First Column (e.g., A)", description = "The column letter used for the first column to concatenate.")
            @NotEmpty
            String firstColumn,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Second Column (e.g., B)", description = "The column letter used for the second column to concatenate.")
            @NotEmpty
            String secondColumn,

            @Idx(index = "5", type = TEXT)
            @Pkg(label = "Concatenation Result Column (e.g., C)", description = "The column letter used for special character deletion.")
            @NotEmpty
            String outputColumn


    ) {

        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {

            Sheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = sheet.getLastRowNum();
            int startColumn = ExcelUtils.columnLetterToIndex(firstColumn);
            int endColumn = ExcelUtils.columnLetterToIndex(secondColumn);
            int targetColumn = ExcelUtils.columnLetterToIndex(outputColumn);
            if (startColumn == endColumn) {
                throw new BotCommandException("First and second columns cannot be the same.");
            } else if (targetColumn == startColumn || targetColumn == endColumn) {
                throw new BotCommandException("Output column cannot be the same as an input column.");
            }
            for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                Cell cell = row.getCell(startColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//first part of the concatenated String
                Cell cell2 = row.getCell(endColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//second part of the concatenated String
                Cell cell3 = row.getCell(targetColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//concatenated String
                String val, val2, finalVal;
                if (cell.getCellType() == CellType.NUMERIC) {
                    val = String.valueOf(cell.getNumericCellValue());
                    val2 = String.valueOf(cell2.getNumericCellValue());
                } else {
                    val = cell.getStringCellValue();
                    val2 = cell2.getStringCellValue();
                }
                finalVal = val + val2;
                cell3.setCellValue(finalVal);
            }
            try (FileOutputStream fos = new FileOutputStream(inputFilePath)) {
                workbook.write(fos);
            }
            return new StringValue("Output Saved as: " + inputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }
}
