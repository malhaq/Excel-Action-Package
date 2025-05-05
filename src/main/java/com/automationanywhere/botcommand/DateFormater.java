package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.ParseException;
import java.util.*;
import java.text.SimpleDateFormat;


import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "DateFormater", label = "Formate the date column", description = "Format the date column to the chosen date format", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "Date Formating status", return_type = STRING, return_required = true)

public class DateFormater {

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
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "yyyy-MM-dd", value = "yyyy-MM-dd")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "MM/dd/yyyy", value = "MM/dd/yyyy")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "dd/MM/yyyy", value = "dd/MM/yyyy")),
                    @Idx.Option(index = "3.4", pkg = @Pkg(label = "MMMM d, yyyy", value = "MMMM d, yyyy")),
                    @Idx.Option(index = "3.5", pkg = @Pkg(label = "MMM dd, yyyy", value = "MMM dd, yyyy")),
                    @Idx.Option(index = "3.6", pkg = @Pkg(label = "EEEE, MMM d", value = "EEEE, MMM d")),
                    @Idx.Option(index = "3.7", pkg = @Pkg(label = "yyyy/MM/dd", value = "yyyy/MM/dd")),
                    @Idx.Option(index = "3.8", pkg = @Pkg(label = "dd-MM-yyyy", value = "dd-MM-yyyy")),
                    @Idx.Option(index = "3.9", pkg = @Pkg(label = "MM-dd-yyyy", value = "MM-dd-yyyy"))

            })
            @Pkg(label = "Date Format: ")
            @NotEmpty
            String format,

            @Idx(index = "4", type = BOOLEAN)
            @Pkg(label = "With Headers",description = "If your table has headers set this option to \"True\" if not set it to \"False\"")
            @NotEmpty
            Boolean headers


    ) {
        try {
            int targetColumn = ExcelUtils.columnLetterToIndex(column);
            formatColumn(inputFilePath, targetColumn, format, headers);
            return new StringValue("completed successfully");
        } catch (Exception e) {
            throw new BotCommandException("Error formatting dates: " + e.getMessage());
        }
    }

    private void formatColumn(String inputFile, int targetColumn, String format, Boolean headers) throws Exception {
        File file = new File(inputFile);
        File backupFile = createBackupFile(file);
        File tempFile = File.createTempFile("excel_date", ".xlsx");

        try (Workbook workbook = WorkbookFactory.create(file);
             FileOutputStream out = new FileOutputStream(tempFile)) {

            Sheet sheet = workbook.getSheetAt(0);

            if (sheet.getPhysicalNumberOfRows() == 0) {
                throw new BotCommandException("The sheet is empty");
            }
            if (sheet.getRow(0).getLastCellNum() <= targetColumn) {
                throw new BotCommandException("Target column " + targetColumn + " doesn't exist in the sheet");
            }

            int startRow = headers ? 1 : 0;
            SimpleDateFormat outputFormat = new SimpleDateFormat(format);

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(targetColumn);
                    if (cell != null) {
                        Object cellValue = getCellValue(cell);
                        Date dateValue = convertToDate(cellValue);

                        if (dateValue != null) {
                            String formattedDate = outputFormat.format(dateValue);
                            cell.setCellValue(formattedDate);
                        }
                    }
                }
            }

            workbook.write(out);

            Files.copy(tempFile.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);

        } catch (Exception e) {
            Files.copy(backupFile.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);
            throw new BotCommandException("Error formatting dates: " + e.getMessage());
        } finally {
            if (backupFile.exists()) backupFile.delete();
            if (tempFile.exists()) tempFile.delete();
        }
    }

    private Date convertToDate(Object value) {
        System.out.println("HI");

        if (value == null) return null;

        if (value instanceof Date) {
            return (Date) value;
        } else if (value instanceof String) {
            String str = ((String) value).trim();
            if (str.isEmpty()) return null;

            String[] dateFormats = {
                    "MM/dd/yyyy",
                    "M/d/yyyy",
                    "dd/MM/yyyy",
                    "d/M/yyyy",
                    "yyyy-MM-dd",
                    "MM-dd-yyyy",
                    "M-d-yyyy",
                    "dd-MM-yyyy",
                    "d-M-yyyy",
                    "yyyy/MM/dd",
                    "yyyyMMdd",
                    "MMddyyyy",
                    "MMMM d, yyyy",
                    "MMM dd, yyyy",
                    "EEEE, MMM d",
                    "dd-MMM-yyyy",
                    "dd MMM yyyy",
                    "MMM dd, yyyy",
                    "dd-MMM-yy",
                    "yyyy-MMM-dd",
            };

            for (String format : dateFormats) {
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(format);
                    sdf.setLenient(false); // Strict parsing
                    return sdf.parse(str);
                } catch (ParseException e) {
                    // Try next format
                }
            }

            // Try Excel serial date (numeric string)
            try {
                double numericValue = Double.parseDouble(str);
                if (numericValue > 0) {
                    return DateUtil.getJavaDate(numericValue);
                }
            } catch (NumberFormatException e) {
                // Not a numeric string
            }

            // Try SQL date format as last resort
            try {
                return java.sql.Date.valueOf(str);
            } catch (IllegalArgumentException e) {
                throw new BotCommandException("Unable to parse date: " + str +
                        ". Supported formats: MM/dd/yyyy, dd/MM/yyyy, yyyy-MM-dd, etc.");
            }
        }else if (value instanceof Number) {
            // Handle numeric values (Excel serial dates)
            return DateUtil.getJavaDate(((Number) value).doubleValue());
        }

        return null;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            default:
                return null;
        }
    }

    private File createBackupFile(File originalFile) throws IOException {
        File backupFile = new File(originalFile.getParent(), originalFile.getName() + ".bak");
        Files.copy(originalFile.toPath(), backupFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        return backupFile;
    }


}



