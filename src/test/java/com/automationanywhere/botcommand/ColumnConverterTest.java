package com.automationanywhere.botcommand;


import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ColumnConverterTest {

    private ColumnConverter columnConverter;

    @BeforeEach
    void setUp() {
        columnConverter = new ColumnConverter();
    }

    @Test
    void testValidStringToNumberConversion() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output.xlsx";
        String columnRange = "A,C";
        String conversionType = "StringToNumber";

        assertDoesNotThrow(() -> columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testValidNumberToStringConversion() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output1.xlsx";
        String columnRange = "B,D";
        String conversionType = "NumberToString";

        assertDoesNotThrow(() -> columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testInvalidColumnRangeFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output2.xlsx";
        String columnRange = "A"; // Invalid format
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType);
        });


        assertEquals("Invalid column range format. Use A,E format.", exception.getMessage());
    }

    @Test
    void testInvalidColumnLetters() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output3.xlsx";
        String columnRange = "1,E"; // Invalid column
        String conversionType = "NumberToString";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType);
        });

        assertEquals("Invalid column letters: 1", exception.getMessage());
    }

    @Test
    void testStartColumnGreaterThanEndColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output4.xlsx";
        String columnRange = "E,A"; // Start column is greater than end column
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
           columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType);
        });

        assertEquals("Start column cannot be greater than end column.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "test_files/non_existent.xlsx";
        String outputFilePath = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Desktop\\ExcelColumnConvert\\src\\test\\resources\\test_files\\output5.xlsx";
        String columnRange = "A,B";
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testValidStringToNumberConversion2() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Column_Convert/output10.xlsx";
        String columnRange = "B,D";
        String conversionType = "StringToNumber";

        assertDoesNotThrow(() -> columnConverter.action(inputFilePath, outputFilePath, columnRange, conversionType));
        assertTrue(new File(outputFilePath).exists());
    }
}

