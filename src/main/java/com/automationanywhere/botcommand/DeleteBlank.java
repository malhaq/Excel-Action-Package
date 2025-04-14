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
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "BlankRemoval", label = "[[BlankRemoval.label]]", node_label = "[[BlankRemoval.node_label]]", description = "[[BlankRemoval.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "Status Message", return_type = STRING, return_required = true)

public class DeleteBlank {
    private static final Logger LOGGER = Logger.getLogger(DeleteBlank.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(@Idx(index = "1", type = FILE) @Pkg(label = "Input Excel File Path") @NotEmpty String inputFilePath,

//                                @Idx(index = "2", type = TEXT) @Pkg(label = "Output Excel File Path") @NotEmpty String outputFilePath,

                                @Idx(index = "2", type = TEXT) @Pkg(label = "Specified Column (e.g., A)") @NotEmpty String column

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

                if (cell.getCellType() == CellType.BLANK || cell.toString().trim().isEmpty()) {
                    sheet.removeRow(row);
                    if (rowNum < endRow) {
                        sheet.shiftRows(rowNum + 1, endRow, -1);
                        endRow--;
                        rowNum--;
                    }
                }

            }

            try (FileOutputStream fos = new FileOutputStream(inputFilePath)) {
                workbook.write(fos);
            }

            return new StringValue("blank cell detection completed all duplicates have been deleted. Output Saved as: " + inputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }


}
