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
        name = "RemoveSpecialCharacters", label = "[[RemoveSpecialCharacters.label]]", node_label = "[[RemoveSpecialCharacters.node_label]]", description = "[[RemoveSpecialCharacters.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[RemoveSpecialCharacters.return_label]]", return_type = STRING, return_required = true)

public class RemoveSpecialCharacters {
    private static final Logger LOGGER = Logger.getLogger(RemoveSpecialCharacters.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input Excel File Path", description = "Specify the full path of the Excel file to be modified.")
            @NotEmpty
            String inputFilePath,


            @Idx(index = "2", type = FILE)
            @Pkg(label = "Output Excel File Path", description = "Path where the modified Excel file will be saved.")
            @NotEmpty
            String outputFilePath,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Target Column", description = "Enter the column letter (e.g., A, B, C) from which special characters should be removed.")
            @NotEmpty
            String column,

            @Idx(index = "4", type = TEXT)
            @Pkg(label = "Special Character", description = "Specify the special character (e.g., -, \", ', /, @, etc.) to be removed.")
            @NotEmpty
            String specialCharacter


    ) {

        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {
            specialCharacter = specialCharacter.trim();
            if (specialCharacter.length() > 1 ) {
                throw new BotCommandException("Invalid input: Enter exactly one special character (e.g., -, @, /, #, etc.). Multiple characters are not allowed.");
            }
            if(Character.isLetter(specialCharacter.charAt(0))){
                throw new BotCommandException("Invalid input: Spacial character cannot be a letter.");
            }
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
//                    System.out.println(val);
                }
                val = val.replace(specialCharacter, "");
                cell.setCellValue(val);
//                System.out.println("");

            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            return new StringValue("Output Saved as: " + outputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }


}
