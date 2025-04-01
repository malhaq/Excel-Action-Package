package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;

public class ExcelUtils {

    // Prevent instantiation of utility class
    private ExcelUtils() {
    }

    /**
     * Converts a column letter to its 0-based index (A=0, B=1, ..., Z=25, AA=26).
     */
    public static int columnLetterToIndex(String letters) {
        if (letters == null || !letters.matches("[A-Z]+")) {
            throw new BotCommandException("Invalid column letters: " + letters);
        }

        letters = letters.toUpperCase(); // Ensure uppercase
        int index = 0;

        for (char c : letters.toCharArray()) {
            if (c < 'A' || c > 'Z') {
                throw new BotCommandException("Invalid column letter: " + letters);
            }
            index = index * 26 + (c - 'A' + 1);
        }

        return index - 1;  // Convert to 0-based index
    }

    /**
     * open the excel file
     *
     * @param inputFilePath
     * @return instance of woorkbook
     * @throws IOException
     */
    public static Workbook openWorkbook(String inputFilePath) throws IOException {
        File file = new File(inputFilePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + inputFilePath);
        }

        FileInputStream fis = new FileInputStream(file);
        return new XSSFWorkbook(fis);
    }

    /**
     * delete a row from the sheet
     *
     * @param sheet
     * @param rowNum intended row to delete
     * @param endRow last row with data
     * @return boolean
     */
    public static boolean deleteRow(Sheet sheet, int rowNum, int endRow) {
        Row row = sheet.getRow(rowNum);
        if (row != null) {
            sheet.removeRow(row);
            if (rowNum < endRow) {
                sheet.shiftRows(rowNum + 1, endRow, -1);
            }
            return true;
        }
        return false;
    }

    /**
     * convert excel file from xls to xlsx
     *
     * @param inputFilePath
     * @param outputFilePath
     * @throws IOException
     */
    public static void convertXlsToXlsx(String inputFilePath, String outputFilePath) throws IOException {
        FileInputStream fis = new FileInputStream(new File(inputFilePath));
        Workbook workbook = new HSSFWorkbook(fis);
        Workbook newWorkBook = new XSSFWorkbook();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet inputSheet = workbook.getSheetAt(i);
            Sheet outputSheet = newWorkBook.createSheet(inputSheet.getSheetName());


            for (Row row : inputSheet) {
                Row newRow = outputSheet.createRow(row.getRowNum());
                for (Cell cell : row) {
                    Cell newCell = newRow.createCell(cell.getColumnIndex(), cell.getCellType());

                    switch (cell.getCellType()) {
                        case STRING:
                            newCell.setCellValue(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            newCell.setCellValue(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            newCell.setCellValue(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            newCell.setCellFormula(cell.getCellFormula());
                            break;
                        default:
                            newCell.setCellValue("");
                    }
                }
            }
        }

        FileOutputStream fos = new FileOutputStream(new File(ensureCorrectExtension(outputFilePath,".xlsx")));
        newWorkBook.write(fos);
        fos.close();
        workbook.close();
        newWorkBook.close();

    }

    /**
     * convert excel file from xlsx to csv
     *
     * @param inputFilePath
     * @param outputFilePath
     * @throws IOException
     */
    public static void convertXlsxToCsv(String inputFilePath, String outputFilePath) throws IOException {
        FileInputStream fis = new FileInputStream(new File(inputFilePath));
        Workbook workbook = new XSSFWorkbook(fis);
        convertLogic(outputFilePath, workbook);

    }

    private static void convertLogic(String outputFilePath, Workbook workbook) throws IOException {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet inputSheet = workbook.getSheetAt(i);
            String sheetName = inputSheet.getSheetName();
            String csvpath = outputFilePath.replaceAll("\\.[^.]+$", "") + i + ".csv";

            FileWriter fw = new FileWriter(csvpath);

            for (Row row : inputSheet) {
                StringBuilder sb = new StringBuilder();
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            sb.append(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            sb.append(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            sb.append(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            CellValue cellValue = evaluator.evaluate(cell);
                            switch (cellValue.getCellType()) {
                                case STRING:
                                    sb.append(cellValue.getStringValue());
                                    break;
                                case NUMERIC:
                                    sb.append(cellValue.getNumberValue());
                                    break;
                                case BOOLEAN:
                                    sb.append(cellValue.getBooleanValue());
                                    break;
                                default:
                                    sb.append("");
                            }
                            break;
                        default:
                            sb.append("");
                    }
                    sb.append(",");
                }
                fw.write(sb.toString().replaceAll(",$", "") + "\n");
            }
            fw.close();
        }
        workbook.close();
    }

    /**
     * convert excel file from xls to csv
     *
     * @param inputFilePath
     * @param outputFilePath
     * @throws IOException
     */
    public static void convertXlsToCsv(String inputFilePath, String outputFilePath) throws IOException {

        FileInputStream fis = new FileInputStream(new File(inputFilePath));
        Workbook workbook = new HSSFWorkbook(fis);
        convertLogic(outputFilePath, workbook);
    }

    /**
     * convert excel file from csv to xlsx
     *
     * @param inputFilePath
     * @param outputFilePath
     * @throws IOException
     */

    public static void convertCsvToXlsx(String inputFilePath, String outputFilePath) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1"); // Create a single sheet named "Sheet1"

        BufferedReader br = new BufferedReader(new FileReader(inputFilePath));
        String line;
        int rowNum = 0;

        while ((line = br.readLine()) != null) {
            Row row = sheet.createRow(rowNum++);
            String[] values = line.split(",");

            for (int i = 0; i < values.length; i++) {
                row.createCell(i).setCellValue(values[i]);
            }
        }
        br.close();

        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        fos.close();
        workbook.close();

    }

    public static String ensureCorrectExtension(String filename, String extension) {
        return filename.replaceAll("\\.[^.]+$", "") + extension;
    }

}


