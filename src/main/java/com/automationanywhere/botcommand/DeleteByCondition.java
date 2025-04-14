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
        name = "DeleteByCondition", label = "[[DeleteByCondition.label]]", node_label = "[[DeleteByCondition.node_label]]", description = "[[DeleteByCondition.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[DeleteByCondition.return_label]]", return_type = STRING, return_required = true)

public class DeleteByCondition {
    private static final Logger LOGGER = Logger.getLogger(DeleteByCondition.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input Excel File Path")
            @NotEmpty
            String inputFilePath,


//            @Idx(index = "2", type = TEXT)
//            @Pkg(label = "Output Excel File Path")
//            @NotEmpty
//            String outputFilePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "Specified Column (e.g., A)")
            @NotEmpty
            String column,

            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Equals (=)", node_label = "Delete where cell value equals {{criteriaValue}}", value = "Equals to =")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Doesn't equal (!=)", node_label = "Delete where cell value does not equal {{criteriaValue}}", value = "Doesn't equals !=")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "Contains", node_label = "Delete where cell value contains {{criteriaValue}}", value = "Contain")),
                    @Idx.Option(index = "3.4", pkg = @Pkg(label = "Doesn't contain", node_label = "Delete where cell value does not contain {{criteriaValue}}", value = "Doesn't Contain"))
            })
            @Pkg(label = "Condition")
            @NotEmpty
            String Condition,

            @Idx(index = "4", type = TEXT)
            @Pkg(label = "Criteria Value ")
            @NotEmpty
            String criteriaValue


    ) {

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
                switch (Condition) {
                    case "Equals to =":
                        if (val.equals(criteriaValue)) {
                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                endRow--;
                                rowNum--;
                            }
                        }
                        break;
                    case "Doesn't equals !=":
                        if (!val.equals(criteriaValue)) {
                            LOGGER.log(Level.INFO, " row {0} ", rowNum + 1);
                            LOGGER.log(Level.INFO, " "+val);
                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                endRow--;
                                rowNum--;
                            }
                        }
                        break;
                    case "Contain":
                        if (val.contains(criteriaValue)) {
                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                endRow--;
                                rowNum--;
                            }
                        }
                        break;
                    case "Doesn't Contain":
                        if (!val.contains(criteriaValue)) {
                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                endRow--;
                                rowNum--;
                            }
                        }
                        break;
                }
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

