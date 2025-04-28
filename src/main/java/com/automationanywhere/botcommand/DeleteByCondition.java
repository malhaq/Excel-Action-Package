package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;


import java.io.*;
import java.util.*;
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
            String criteriaValue,

            @Idx(index = "5", type = BOOLEAN)
            @Pkg(label = "Add another condition")
            @NotNull
            Boolean addNewCondition,


            @Idx(index = "6", type = BOOLEAN)
            @Pkg(label = "AND, OR operator")
            @NotNull
            Boolean logicalOperator,


            @Idx(index = "7", type = TEXT)
            @Pkg(label = "Specified Column (e.g., A)")
            @NotEmpty
            String column2,

            @Idx(index = "8", type = SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Equals (=)", node_label = "Delete where cell value equals {{criteriaValue2}}", value = "Equals to =")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "Doesn't equal (!=)", node_label = "Delete where cell value does not equal {{criteriaValue2}}", value = "Doesn't equals !=")),
                    @Idx.Option(index = "8.3", pkg = @Pkg(label = "Contains", node_label = "Delete where cell value contains {{criteriaValue2}}", value = "Contain")),
                    @Idx.Option(index = "8.4", pkg = @Pkg(label = "Doesn't contain", node_label = "Delete where cell value does not contain {{criteriaValue2}}", value = "Doesn't Contain"))
            })
            @Pkg(label = "Condition")
            @NotEmpty
            String Condition2,

            @Idx(index = "9", type = TEXT)
            @Pkg(label = "Criteria Value ")
            @NotEmpty
            String criteriaValue2


    ) {
        if (inputFilePath == null || inputFilePath.isEmpty()) {
            throw new BotCommandException("Input Excel File Path must not be empty.");
        }
        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {

            if (column == null || column.isEmpty()) {
                throw new IllegalArgumentException("Specified Column must not be empty.");
            }

            if (Condition == null || Condition.isEmpty()) {
                throw new IllegalArgumentException("Condition must not be empty.");
            }

            if (criteriaValue == null || criteriaValue.isEmpty()) {
                throw new IllegalArgumentException("Criteria Value must not be empty.");
            }

            // Validate condition options
            List<String> validConditions = Arrays.asList(
                    "Equals to =",
                    "Doesn't equals !=",
                    "Contain",
                    "Doesn't Contain"
            );

            if (!validConditions.contains(Condition)) {
                throw new IllegalArgumentException("Condition must be one of: " + validConditions);
            }

            if (addNewCondition) {
                if (column2 == null || column2.isEmpty()) {
                    throw new IllegalArgumentException("Specified Column must not be empty.");
                }

                if (Condition2 == null || Condition2.isEmpty()) {
                    throw new IllegalArgumentException("Condition must not be empty.");
                }

                if (criteriaValue2 == null || criteriaValue2.isEmpty()) {
                    throw new IllegalArgumentException("Criteria Value must not be empty.");
                }

                if (!validConditions.contains(Condition2)) {
                    throw new IllegalArgumentException("Condition must be one of: " + validConditions);
                }
            }

            Sheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = sheet.getLastRowNum();
            int startColumn = ExcelUtils.columnLetterToIndex(column);

            if (!addNewCondition) {
                switch (Condition) {
                    case "Equals to =":
                        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                            if (equals(sheet, rowNum, startColumn, criteriaValue)) {
                                if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                    endRow--;
                                    rowNum--;
                                }
                            }
                        }
                        break;
                    case "Doesn't equals !=":
                        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                            if (notEqual(sheet, rowNum, startColumn, criteriaValue)) {
                                if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                    endRow--;
                                    rowNum--;
                                }
                            }
                        }
                        break;
                    case "Contain":
                        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                            if (contains(sheet, rowNum, startColumn, criteriaValue)) {
                                if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                    endRow--;
                                    rowNum--;
                                }
                            }
                        }
                        break;
                    case "Doesn't Contain":
                        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                            if (doesNotContain(sheet, rowNum, startColumn, criteriaValue)) {
                                if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                    endRow--;
                                    rowNum--;
                                }
                            }
                        }
                        break;
                    default:
                        throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                }
            } else {

                int startColumn2 = ExcelUtils.columnLetterToIndex(column2);
                if (logicalOperator) {
                    // OR operator
                    switch (Condition) {
                        case "Equals to =":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) || equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) || notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) || contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) || doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Doesn't equals !=":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) || equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) || notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) || contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) || doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Contain":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) || equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) || notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) || contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) || doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Doesn't Contain":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) || equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) || notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) || contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) || doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;
                        default:
                            throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");


                    }


                } else {
                    switch (Condition) {
                        case "Equals to =":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) && equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) && notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) && contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (equals(sheet, rowNum, startColumn, criteriaValue) && doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Doesn't equals !=":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) && equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) && notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) && contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (notEqual(sheet, rowNum, startColumn, criteriaValue) && doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Contain":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) && equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) && notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) && contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (contains(sheet, rowNum, startColumn, criteriaValue) && doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;

                        case "Doesn't Contain":
                            switch (Condition2) {
                                case "Equals to =":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) && equals(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't equals !=":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) && notEqual(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) && contains(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                case "Doesn't Contain":
                                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                                        if (doesNotContain(sheet, rowNum, startColumn, criteriaValue) && doesNotContain(sheet, rowNum, startColumn2, criteriaValue2)) {
                                            if (ExcelUtils.deleteRow(sheet, rowNum, endRow)) {
                                                endRow--;
                                                rowNum--;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");
                            }
                            break;
                        default:
                            throw new RuntimeException("Some thing wrong you have reached the Default value in the switch statement");


                    }
                }
            }
            try (FileOutputStream fos = new FileOutputStream(inputFilePath)) {
                workbook.write(fos);
            }
            return new StringValue("Output Saved as: " + inputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("Error: " + e.getMessage());
        } catch (IllegalArgumentException e) {
            throw new BotCommandException("Error: " + e.getMessage());
        }
    }

    private String getRow(Sheet sheet, int rowNum, int startColumn) {
        Row row = sheet.getRow(rowNum);
        if (row == null) return null;
        Cell cell = row.getCell(startColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        String val;
        if (cell.getCellType() == CellType.NUMERIC) {
            val = String.valueOf(cell.getNumericCellValue());
        } else {
            val = cell.getStringCellValue();
        }
        return val;
    }

    private boolean equals(Sheet sheet, int rowNum, int startColumn, String criteriaValue) {
        return getRow(sheet, rowNum, startColumn).equals(criteriaValue);
    }

    private boolean notEqual(Sheet sheet, int rowNum, int startColumn, String criteriaValue) {
        return !getRow(sheet, rowNum, startColumn).equals(criteriaValue);
    }

    private boolean contains(Sheet sheet, int rowNum, int startColumn, String criteriaValue) {
        return getRow(sheet, rowNum, startColumn).contains(criteriaValue);
    }

    private boolean doesNotContain(Sheet sheet, int rowNum, int startColumn, String criteriaValue) {
        return !getRow(sheet, rowNum, startColumn).contains(criteriaValue);
    }
}

