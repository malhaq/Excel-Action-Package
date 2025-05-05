package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.*;

class SortExcelTest {


    private SortExcel sortExcel;

    @BeforeEach
    void setUp() {
        sortExcel = new SortExcel();
    }


    @Test
    void testSortNumericColumnAscending() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\SortExcel\\test_data.xlsx";

        Value<String> result = sortExcel.action(
                testFile,
                "A",  // Numeric column
                "Numeric",
                "ascend",
                true
        );

        assertEquals("Excel File Sorted Successfully", result.get());

        // Verify the sorting
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(testFile))) {
            Sheet sheet = workbook.getSheetAt(0);

            assertTrue(sheet.getRow(1).getCell(0).getNumericCellValue() < sheet.getRow(2).getCell(0).getNumericCellValue());
            assertTrue(sheet.getRow(2).getCell(0).getNumericCellValue()< sheet.getRow(3).getCell(0).getNumericCellValue());


            int lastRow = sheet.getLastRowNum();
            assertTrue(sheet.getRow(lastRow).getCell(0).getNumericCellValue() > 3.0);
        }
    }

    @Test
    void testSortStringColumnDescending() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\SortExcel\\test_data1.xlsx";

        Value<String> result = sortExcel.action(
                testFile,
                "B",  // String column
                "Alphabet",
                "descend",
                true
        );

        assertEquals("Excel File Sorted Successfully", result.get());

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(testFile))) {
            Sheet sheet = workbook.getSheetAt(0);

            // Check first row after header should start with later letter
            String firstValue = sheet.getRow(1).getCell(1).getStringCellValue();
            String secondValue = sheet.getRow(2).getCell(1).getStringCellValue();
            assertTrue(firstValue.compareToIgnoreCase(secondValue) > 0);
        }
    }

    @Test
    void testSortDateColumnWithNullValues() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\SortExcel\\date_sort .xlsx";

        Value<String> result = sortExcel.action(
                testFile,
                "A",  // Date column
                "Date",
                "ascend",
                true
        );

        assertEquals("Excel File Sorted Successfully", result.get());

    }

    @Test
    void testSortWithoutHeaders() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\SortExcel\\test_no_headers.xlsx";
        System.out.println("I am here");
        Value<String> result = sortExcel.action(
                testFile,
                "A",  // First column
                "Numeric",
                "ascend",
                false  // No headers
        );
        System.out.println("here I am");
        assertEquals("Excel File Sorted Successfully", result.get());

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(testFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("Test with out headers first assertion done");
            System.out.println(sheet.getRow(0).getCell(0).getNumericCellValue());
            System.out.println(sheet.getRow(1).getCell(0).getNumericCellValue());
            double firstValue = sheet.getRow(0).getCell(0).getNumericCellValue();
            double secondValue = sheet.getRow(1).getCell(0).getNumericCellValue();
            assertTrue(firstValue < secondValue);
        }
    }

    @Test
    void testInvalidFilePathThrowsException() {
        assertThrows(BotCommandException.class, () -> {
            sortExcel.action(
                    "nonexistent_file.xlsx",
                    "A",
                    "Numeric",
                    "ascend",
                    true
            );
        });
    }

    @Test
    void testMixedDataTypesInColumn() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\SortExcel\\test_mixed_data.xlsx";

        Value<String> result = sortExcel.action(
                testFile,
                "A",  // Column with mixed numbers and strings
                "Numeric",
                "ascend",
                false
        );

        assertEquals("Excel File Sorted Successfully", result.get());

    }
}