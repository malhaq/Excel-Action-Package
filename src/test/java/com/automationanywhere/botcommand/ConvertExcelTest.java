package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ConvertExcelTest {

    private ConvertExcel convertExcel;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "original/";
    private final String workingDir = testResourcesDir + "working/ConvertExcel/";

    @BeforeEach
    void setUp() {
        convertExcel = new ConvertExcel();
    }

    @Test
    void testXLSXtoCSVConversion() {
        String inputFilePath = workingDir+"valid-xlsx-test-input.xlsx";
        String outputFilePath = workingDir+"valid-xlsx-test-input.csv";
        String conversionType = "XLSXtoCSV";
        assertDoesNotThrow(() -> {
            String result = convertExcel.action(inputFilePath, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testXLStoCSVConversion() {
        String inputFilePath = workingDir+"valid-xls-test-input.xls";
        String outputFilePath = workingDir+"valid-xls-test-input.csv";
        String conversionType = "XLStoCSV";

        assertDoesNotThrow(() -> {
            String result = convertExcel.action(inputFilePath, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testXLStoXLSXConversion() {
        String inputFilePath = workingDir+"valid-xls-test-input 1.xls";
        String outputFilePath = workingDir+"valid-xls-test-input 1.xlsx";
        String conversionType = "XLStoXLSX";

        assertDoesNotThrow(() -> {
            String result = convertExcel.action(inputFilePath, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testCSVtoXLSXConversion() {
        String inputFilePath = workingDir+"valid-csv-test-input.csv";
        String outputFilePath = workingDir+"valid-csv-test-input.xlsx";
        String conversionType = "CSVtoXLSX";

        assertDoesNotThrow(() -> {
            String result = convertExcel.action(inputFilePath, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = workingDir+"non_existent_file.xlsx";
        String outputFilePath = workingDir+"valid-test-input.xlsx";
        String conversionType = "XLSXtoCSV";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            convertExcel.action(inputFilePath, conversionType);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

}