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
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ColumnTypeConvert", label = "[[ConvertColumn.label]]", node_label = "[[ConvertColumn.node_label]]", description = "[[ConvertColumn.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[ConvertColumn.return_label]]", return_type = STRING, return_required = true)
public class ColumnConverter {

    private static final Logger LOGGER = Logger.getLogger(ColumnConverter.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(description = "Provide the full path of the Excel file that contains the data to be converted.",node_label = "Select the Excel file to process",label = "[[ConvertColumn.inputPath.label]]")
            @NotEmpty
            String inputFilePath,

            @Idx(index = "2", type = FILE)
            @Pkg(description = "Provide the full path you want the out put file to be stored in." ,label = "[[ConvertColumn.outputPath.label]]")
            @NotEmpty
            String outputFilePath,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Column Range (e.g., A,E)")
            @NotEmpty
            String columnRange,

            @Idx(index = "4", type = SELECT, options = {
                    @Idx.Option(index = "4.1",pkg = @Pkg(node_label = "[[ConvertColumn.4.1.node_label]]", label = "[[ConvertColumn.4.1.label]]", value = "StringToNumber")),
                    @Idx.Option(index = "4.2",pkg = @Pkg(node_label = "[[ConvertColumn.4.2.node_label]]", label = "[[ConvertColumn.4.2.label]]", value = "NumberToString"))
            })
            @Pkg(label = "Conversion Type")
            @NotEmpty
            String conversionType
    ) {

        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {

            Sheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = sheet.getLastRowNum();
            int[] columnIndexes = parseColumnRange(columnRange);
            int startColumn = columnIndexes[0];
            int endColumn = columnIndexes[1];


            for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                for (int colNum = startColumn; colNum <= endColumn; colNum++) {
                    Cell cell = row.getCell(colNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    try {
                        if ("StringToNumber".equals(conversionType)) {
                            // Convert String to Number
                            if (cell.getCellType() == CellType.STRING) {
                                String text = cell.getStringCellValue().trim();
                                if (text.matches("-?\\d+(\\.\\d+)?")) {
                                    double num = Double.parseDouble(text);
                                    cell.setCellValue(num);
                                }
                            }
                        } else if ("NumberToString".equals(conversionType)) {
                            // Convert Number to String
                            if (cell.getCellType() == CellType.NUMERIC) {
                                cell.setCellValue(String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    } catch (Exception e) {
                        LOGGER.log(Level.WARNING, "Error processing row " + rowNum + ", column " + colNum, e);
                        throw new BotCommandException("Error processing with message: " + e.getMessage());
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            return new StringValue("Conversion completed. Saved as: " + outputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }
    /**
     * Converts column letter range  to numerical indexes (0-based).
     */
    private int[] parseColumnRange(String range) throws BotCommandException {
        String[] parts = range.split(",");
        if (parts.length != 2) {
            throw new BotCommandException("Invalid column range format. Use A,E format.");
        }
        int startCol = ExcelUtils.columnLetterToIndex(parts[0]);
        int endCol = ExcelUtils.columnLetterToIndex(parts[1]);

        if (startCol > endCol) {
            throw new BotCommandException("Start column cannot be greater than end column.");
        }

        return new int[]{startCol, endCol};
    }



}