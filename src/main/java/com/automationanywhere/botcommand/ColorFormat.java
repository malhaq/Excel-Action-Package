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
import java.io.ByteArrayOutputStream;
import java.io.*;
import java.util.logging.Logger;
import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ColorFormat", label = "[[ColorFormate.label]]", node_label = "[[ColorFormate.node_label]]", description = "[[ColorFormate.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[ColorFormate.return_label]]", return_type = STRING, return_required = true)


public class ColorFormat {

    private static final Logger LOGGER = Logger.getLogger(DeleteByCondition.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "File path of the input Excel file")
            @NotEmpty
            String inputFilePath,

//            @Idx(index = "2", type = FILE)
//            @Pkg(label = "File path to save the modified Excel file")
//            @NotEmpty
//            String outputFilePath,

            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "By Column", node_label = "Color the cells in {{range}}", value = "Column")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "By Row", node_label = "Color all cells in row{{range}}", value = "Row"))
            })
            @Pkg(label = "Select whether to color by column or row")
            @NotEmpty
            String Condition,

            @Idx(index = "3", type = TEXT)
            @Pkg(description="(By Column: e.g. A1:C5, A1, A) (By Row: only integer values e.g. 1:14, 5)",label = "Range")
            @NotEmpty
            String range,

            @Idx(index = "4", type = SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "BLACK", value = "8")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "WHITE", value = "9")),
                    @Idx.Option(index = "4.3", pkg = @Pkg(label = "RED", value = "10")),
                    @Idx.Option(index = "4.4", pkg = @Pkg(label = "BRIGHT_GREEN", value = "11")),
                    @Idx.Option(index = "4.5", pkg = @Pkg(label = "BLUE", value = "12")),
                    @Idx.Option(index = "4.6", pkg = @Pkg(label = "YELLOW", value = "13")),
                    @Idx.Option(index = "4.7", pkg = @Pkg(label = "PINK", value = "14")),
                    @Idx.Option(index = "4.8", pkg = @Pkg(label = "TURQUOISE", value = "15")),
                    @Idx.Option(index = "4.9", pkg = @Pkg(label = "DARK_RED", value = "16")),
                    @Idx.Option(index = "4.10", pkg = @Pkg(label = "GREEN", value = "17")),
                    @Idx.Option(index = "4.11", pkg = @Pkg(label = "DARK_BLUE", value = "18")),
                    @Idx.Option(index = "4.12", pkg = @Pkg(label = "DARK_YELLOW", value = "19")),
                    @Idx.Option(index = "4.13", pkg = @Pkg(label = "VIOLET", value = "20")),
                    @Idx.Option(index = "4.14", pkg = @Pkg(label = "TEAL", value = "21")),
                    @Idx.Option(index = "4.15", pkg = @Pkg(label = "GREY_25_PERCENT", value = "22")),
                    @Idx.Option(index = "4.16", pkg = @Pkg(label = "GREY_50_PERCENT", value = "23")),
                    @Idx.Option(index = "4.17", pkg = @Pkg(label = "CORNFLOWER_BLUE", value = "24")),
                    @Idx.Option(index = "4.18", pkg = @Pkg(label = "MAROON", value = "25")),
                    @Idx.Option(index = "4.19", pkg = @Pkg(label = "LEMON_CHIFFON", value = "26")),
                    @Idx.Option(index = "4.20", pkg = @Pkg(label = "LIGHT_TURQUOISE1", value = "27")),
                    @Idx.Option(index = "4.21", pkg = @Pkg(label = "ORCHID", value = "28")),
                    @Idx.Option(index = "4.22", pkg = @Pkg(label = "CORAL", value = "29")),
                    @Idx.Option(index = "4.23", pkg = @Pkg(label = "ROYAL_BLUE", value = "30")),
                    @Idx.Option(index = "4.24", pkg = @Pkg(label = "LIGHT_CORNFLOWER_BLUE", value = "31")),
                    @Idx.Option(index = "4.25", pkg = @Pkg(label = "SKY_BLUE", value = "40")),
                    @Idx.Option(index = "4.26", pkg = @Pkg(label = "LIGHT_TURQUOISE", value = "41")),
                    @Idx.Option(index = "4.27", pkg = @Pkg(label = "LIGHT_GREEN", value = "42")),
                    @Idx.Option(index = "4.28", pkg = @Pkg(label = "LIGHT_YELLOW", value = "43")),
                    @Idx.Option(index = "4.29", pkg = @Pkg(label = "PALE_BLUE", value = "44")),
                    @Idx.Option(index = "4.30", pkg = @Pkg(label = "ROSE", value = "45")),
                    @Idx.Option(index = "4.31", pkg = @Pkg(label = "LAVENDER", value = "46")),
                    @Idx.Option(index = "4.32", pkg = @Pkg(label = "TAN", value = "47")),
                    @Idx.Option(index = "4.33", pkg = @Pkg(label = "LIGHT_BLUE", value = "48")),
                    @Idx.Option(index = "4.34", pkg = @Pkg(label = "AQUA", value = "49")),
                    @Idx.Option(index = "4.35", pkg = @Pkg(label = "LIME", value = "50")),
                    @Idx.Option(index = "4.36", pkg = @Pkg(label = "GOLD", value = "51")),
                    @Idx.Option(index = "4.37", pkg = @Pkg(label = "LIGHT_ORANGE", value = "52")),
                    @Idx.Option(index = "4.38", pkg = @Pkg(label = "ORANGE", value = "53")),
                    @Idx.Option(index = "4.39", pkg = @Pkg(label = "BLUE_GREY", value = "54")),
                    @Idx.Option(index = "4.40", pkg = @Pkg(label = "GREY_40_PERCENT", value = "55")),
                    @Idx.Option(index = "4.41", pkg = @Pkg(label = "DARK_TEAL", value = "56")),
                    @Idx.Option(index = "4.42", pkg = @Pkg(label = "SEA_GREEN", value = "57")),
                    @Idx.Option(index = "4.43", pkg = @Pkg(label = "DARK_GREEN", value = "58")),
                    @Idx.Option(index = "4.44", pkg = @Pkg(label = "OLIVE_GREEN", value = "59")),
                    @Idx.Option(index = "4.45", pkg = @Pkg(label = "BROWN", value = "60")),
                    @Idx.Option(index = "4.46", pkg = @Pkg(label = "PLUM", value = "61")),
                    @Idx.Option(index = "4.47", pkg = @Pkg(label = "INDIGO", value = "62")),
                    @Idx.Option(index = "4.48", pkg = @Pkg(label = "GREY_80_PERCENT", value = "63")),
            })
            @Pkg(label = "Select a color for highlighting the cells")
            @NotEmpty
            String colorIndx




    ) {
        boolean single;
        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {
            single = !range.contains(":");
            Sheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = sheet.getLastRowNum();
            int startColumn = 0;
            int endColumn;
            range = range.trim();

            //creating the style instance to apply it on cells
            CellStyle style = workbook.createCellStyle();
            //set the highlight color to the style instance
            style.setFillForegroundColor((short)Integer.parseInt(colorIndx));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //set the boarder to the highlight style instance
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);



            if (single) {
                if (Condition.equals( "Column")) {
                    if (range.matches(".*\\d.*")) {
                        int digitIndx = ExcelUtils.firstDigitIndex(range);
                        startColumn = ExcelUtils.columnLetterToIndex(range.substring(0, digitIndx));
                        startRow = Integer.parseInt(range.substring(digitIndx).trim())-1;
                    }else {
                        startColumn = ExcelUtils.columnLetterToIndex(range);
                    }
                    for(int i = startRow; i <= endRow; i++) {
                        Row row = sheet.getRow(i);
                        if (row == null) continue;
                        Cell cell = row.getCell(startColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        cell.setCellStyle(style);
                    }
                } else {
                    if(!range.matches("\\d+")){
                        throw new BotCommandException("Row value must be a number or a range of numbers.");
                    }
                    startRow = Integer.parseInt(range)-1;
                    Row row = sheet.getRow(startRow);
                    endColumn = row.getLastCellNum();
                    for(int i = startColumn; i <= endColumn; i++) {
                        Cell cell = row.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        cell.setCellStyle(style);
                    }
                }

            }else {
                String [] array = range.split(":");
                System.out.println(array[0]);
                System.out.println(array[1]);
                if (Condition.equals("Row")) {
                    startRow = Integer.parseInt(array[0])-1;
                    endRow = Integer.parseInt(array[1])-1;
                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        endColumn = row.getLastCellNum();
                        for(int i = startColumn; i < endColumn; i++) {
                            Cell cell = row.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cell.setCellStyle(style);
                        }
                    }
                } else {
                    if (array[0].matches(".*\\d.*")) {
                        int digitIndx = firstDigitIndex(array[0]);
                        System.out.println(array[0].substring(0, digitIndx));
                        System.out.println(array[0].substring(digitIndx).trim());
                        String columnLetter = (array[0] != null) ? array[0].substring(0, digitIndx) : null;
                        System.out.println(columnLetter);
                        startColumn = ExcelUtils.columnLetterToIndex(columnLetter);
                        startRow = Integer.parseInt(array[0].substring(digitIndx).trim())-1;
                    }else {
                        startColumn = ExcelUtils.columnLetterToIndex(array[0]);
                    }
                    if (array[1].matches(".*\\d.*")) {
                        int digitIndx = firstDigitIndex(array[1]);
//                        String columnLetter = (array[0] != null) ? array[1].substring(0, digitIndx) : null;
                        endColumn = ExcelUtils.columnLetterToIndex(array[1].substring(0, digitIndx));
                        endRow = Integer.parseInt(array[1].substring(digitIndx).trim())-1;
                    }else {
                        endColumn = ExcelUtils.columnLetterToIndex(array[1]);
                    }
                    for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        for(int i = startColumn; i <= endColumn; i++) {
                            Cell cell = row.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cell.setCellStyle(style);
                        }
                    }

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
