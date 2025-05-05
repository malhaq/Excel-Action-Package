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
import java.util.*;
import java.text.SimpleDateFormat;





import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "RowSort", label = "Sort Excel Sheet rows", description = "Sort the data in an excel sheet according to one column in an ascending or descending order", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[DeleteByCondition.return_label]]", return_type = STRING, return_required = true)

public class SortExcel {

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
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Numbers", value = "Numeric")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Alphabetical Order", value = "Alphabet")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "Date", value = "Date"))
            })
            @Pkg(label = "Sort by:")
            @NotEmpty
            String sortBy,

            @Idx(index = "4", type = SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Ascending", value = "ascend")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Descending", value = "descend"))
            })
            @Pkg(label = "Sort order:")
            @NotEmpty
            String order,

            @Idx(index = "5", type = BOOLEAN)
            @Pkg(label = "With Headers")
            @NotEmpty
            Boolean headers


    ) {
        try {

            int targetColumn = ExcelUtils.columnLetterToIndex(column);
            boolean ascending = "ascend".equals(order);
            sortExcelFile(inputFilePath, targetColumn, sortBy, ascending, headers);
            return new StringValue("Excel File Sorted Successfully");

        } catch (Exception e) {
            System.out.println(e.getMessage());
            throw new BotCommandException("Some thing wrong: " + e.getMessage());
        }

    }

    private void sortExcelFile(String inputFile, int targetColumn, String sortBy, boolean ascending, boolean headers) throws IOException {
        File file = new File(inputFile);
        File backupFile = createBackupFile(file);
        File tempFile = File.createTempFile("excel_sort_temp", ".xlsx");

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Validate sheet and column
            if (sheet.getPhysicalNumberOfRows() == 0) {
                throw new BotCommandException("The sheet is empty");
            }
            if (sheet.getRow(0).getLastCellNum() <= targetColumn) {
                throw new BotCommandException("Target column " + targetColumn + " doesn't exist in the sheet");
            }

            // Collect rows and their data
            int startRow = headers ? 1 : 0;
            List<List<Object>> rowData = new ArrayList<>();

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    List<Object> data = new ArrayList<>();
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        data.add(getCellValue(row.getCell(j)));
                    }
                    rowData.add(data);
                }
            }

            // Create and apply comparator
            Comparator<List<Object>> comparator = createComparator(targetColumn, sortBy);
            if (!ascending) {
                comparator = comparator.reversed();
            }
            rowData.sort(comparator);

            // Clear and rebuild sheet
            int lastRowNum = sheet.getLastRowNum();
            for (int i = startRow; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }

            // Rebuild rows with preserved formatting
            for (int i = 0; i < rowData.size(); i++) {
                Row newRow = sheet.createRow(headers ? i + 1 : i);
                List<Object> data = rowData.get(i);
                for (int j = 0; j < data.size(); j++) {
                    Cell newCell = newRow.createCell(j);
                    setCellValue(newCell, data.get(j), workbook);
                }
            }

            // Write to temp file
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
            }

            // Replace original file
            Files.move(tempFile.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);
        } catch (Exception e) {
            // Restore backup if error occurs
            try {
                Files.move(backupFile.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException restoreEx) {
                e.addSuppressed(restoreEx);
            }
            throw new IOException("Failed to sort Excel file. Original file has been restored."+ e.getMessage());
        } finally {
            Files.deleteIfExists(tempFile.toPath());
            Files.deleteIfExists(backupFile.toPath());
        }
    }

    private Comparator<List<Object>> createComparator(int columnIndex, String sortBy) {
        return (row1, row2) -> {
            Object val1 = row1.size() > columnIndex ? row1.get(columnIndex) : null;
            Object val2 = row2.size() > columnIndex ? row2.get(columnIndex) : null;

            if (val1 == null && val2 == null) return 0;
            if (val1 == null) return 1;
            if (val2 == null) return -1;

            if (sortBy.equals("Alphabet")) {
                return val1.toString().compareToIgnoreCase(val2.toString());
            }

            try {
                if (sortBy.equals("Numeric")) {
                    Double num1 = convertToNumber(val1);
                    Double num2 = convertToNumber(val2);
                    if (num1 != null && num2 != null) {
                        return num1.compareTo(num2);
                    }
                }
                else if (sortBy.equals("Date")) {
                    Date date1 = convertToDate(val1);
                    Date date2 = convertToDate(val2);
                    if (date1 != null && date2 != null) {
                        return date1.compareTo(date2);
                    }
                }
            } catch (Exception e) {
                throw new RuntimeException("comparator creation Error");
            }


            return compareMixedTypes(val1, val2, sortBy);
        };
    }

    private int compareMixedTypes(Object val1, Object val2, String sortBy) {

        int typeOrder1 = getTypeOrder(val1, sortBy);
        int typeOrder2 = getTypeOrder(val2, sortBy);

        if (typeOrder1 != typeOrder2) {
            return Integer.compare(typeOrder1, typeOrder2);
        }

        switch (sortBy) {
            case "Numeric":
                Double num1 = convertToNumber(val1);
                Double num2 = convertToNumber(val2);
                if (num1 != null && num2 != null) {
                    return num1.compareTo(num2);
                }
                break;
            case "Date":
                Date date1 = convertToDate(val1);
                Date date2 = convertToDate(val2);
                if (date1 != null && date2 != null) {
                    return date1.compareTo(date2);
                }
                break;
        }

        return val1.toString().compareToIgnoreCase(val2.toString());
    }

    private int getTypeOrder(Object value, String sortBy) {
        if (value == null) return 0;

        if (sortBy.equals("Numeric")) {
            if (value instanceof Number) return 1;
            if (convertToNumber(value) != null) return 1; // String that can be parsed as number
            return 2;
        }else if (sortBy.equals("Date")) {
            if (value instanceof Date) return 1;
            if (convertToDate(value) != null) return 1; // String that can be parsed as date
            return 2;
        }

        if (value instanceof Date) return 1;
        if (value instanceof Number) return 2;
        return 3;
    }

    private Double convertToNumber(Object value) {
        if (value == null) return null;
        if (value instanceof Number) return ((Number) value).doubleValue();
        if (value instanceof String) {
            try {
                return Double.parseDouble(value.toString().trim());
            } catch (NumberFormatException e) {
                return null;
            }
        }
        return null;
    }

    private Date convertToDate(Object value) {
        if (value == null) return null;
        if (value instanceof Date) return (Date) value;
        if (value instanceof String) {
            String[] dateFormats = {
                    "yyyy-MM-dd", "MM/dd/yyyy", "dd-MMM-yyyy",
                    "yyyy/MM/dd", "dd.MM.yyyy", "MMM dd, yyyy",
                    "yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy HH:mm:ss"
            };

            for (String format : dateFormats) {
                try {
                    return new SimpleDateFormat(format).parse(value.toString().trim());
                } catch (Exception e) {
                    throw new RuntimeException("date conversion error");
                }
            }
        }
        return null;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN: return cell.getBooleanCellValue();
            case FORMULA: return cell.getCellFormula();
            default: return null;
        }
    }

    private void setCellValue(Cell cell, Object value, Workbook workbook) {
        if (value == null) return;

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
            CellStyle style = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            style.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
            cell.setCellStyle(style);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof String && ((String) value).startsWith("=")) {
            cell.setCellFormula(((String) value).substring(1));
        }
    }

    private File createBackupFile(File originalFile) throws IOException {
        File backupFile = new File(originalFile.getParent(), originalFile.getName() + ".bak");
        Files.copy(originalFile.toPath(), backupFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        return backupFile;
    }

}
