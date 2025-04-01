package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ConvertExcelTest {

    private ConvertExcel convertExcel;

    @BeforeEach
    void setUp() {
        convertExcel = new ConvertExcel();
    }

    @Test
    void testXLSXtoCSVConversion() {
        String inputFilePath = "src/test/resources/test_files/original/valid-test-input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ConvertExcel/xlsxToCsvOutput.csv";
        String conversionType = "XLSXtoCSV";

        assertDoesNotThrow(() -> convertExcel.action(inputFilePath, outputFilePath, conversionType));
        assertTrue(new File(outputFilePath.replace(".csv", "0.csv")).exists());
    }

    @Test
    void testXLStoCSVConversion() {
        String inputFilePath = "src/test/resources/test_files/original/valid-test-input.xls";
        String outputFilePath = "src/test/resources/test_files/ConvertExcel/xlsToCsvOutput.csv";
        String conversionType = "XLStoCSV";

        assertDoesNotThrow(() -> convertExcel.action(inputFilePath, outputFilePath, conversionType));
        assertTrue(new File(outputFilePath.replace(".csv", "0.csv")).exists());
    }

    @Test
    void testXLStoXLSXConversion() {
        String inputFilePath = "src/test/resources/test_files/original/valid-test-input.xls";
        String outputFilePath = "src/test/resources/test_files/ConvertExcel/xlsToXlsxOutput.xlsx";
        String conversionType = "XLStoXLSX";

        assertDoesNotThrow(() -> convertExcel.action(inputFilePath, outputFilePath, conversionType));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testCSVtoXLSXConversion() {
        String inputFilePath = "src/test/resources/test_files/original/valid-test-input.csv";
        String outputFilePath = "src/test/resources/test_files/ConvertExcel/csvToXlsxOutput.xlsx";
        String conversionType = "CSVtoXLSX";

        assertDoesNotThrow(() -> convertExcel.action(inputFilePath, outputFilePath, conversionType));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/non_existent_file.xlsx";
        String outputFilePath = "src/test/resources/test_files/ConvertExcel/noOutput.xlsx";
        String conversionType = "XLSXtoCSV";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            convertExcel.action(inputFilePath, outputFilePath, conversionType);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

}