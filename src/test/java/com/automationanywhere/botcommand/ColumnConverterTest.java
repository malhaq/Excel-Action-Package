package com.automationanywhere.botcommand;


import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import static org.junit.jupiter.api.Assertions.*;

class ColumnConverterTest {
    private ColumnConverter columnConverter;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/Column_Convert/";

    @BeforeEach
    void setUp() throws IOException {
        columnConverter = new ColumnConverter();

        // Create working directory if it doesn't exist
        new File(workingDir).mkdirs();

        // Clean working directory
        File[] workingFiles = new File(workingDir).listFiles();
        if (workingFiles != null) {
            for (File file : workingFiles) {
                file.delete();
            }
        }

        // Copy original test files to working directory
        File[] originalFiles = new File(originalsDir).listFiles();
        if (originalFiles != null) {
            for (File original : originalFiles) {
                Path source = original.toPath();
                Path target = new File(workingDir + original.getName()).toPath();
                Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
            }
        }
    }

    @Test
    void testValidStringToNumberConversion() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output.xlsx";
        String columnRange = "A,C";
        String conversionType = "StringToNumber";

        assertDoesNotThrow(() -> {
            String result = columnConverter.action(inputFilePath, columnRange, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testValidNumberToStringConversion() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output1.xlsx";
        String columnRange = "B,D";
        String conversionType = "NumberToString";

        assertDoesNotThrow(() -> {
            String result = columnConverter.action(inputFilePath, columnRange, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
    }

    @Test
    void testInvalidColumnRangeFormat() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output2.xlsx";
        String columnRange = "A"; // Invalid format
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, columnRange, conversionType);
        });


        assertEquals("Invalid column range format. Use A,E format.", exception.getMessage());
    }

    @Test
    void testInvalidColumnLetters() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output3.xlsx";
        String columnRange = "1,E"; // Invalid column
        String conversionType = "NumberToString";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, columnRange, conversionType);
        });

        assertEquals("Invalid column letters: 1", exception.getMessage());
    }

    @Test
    void testStartColumnGreaterThanEndColumn() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output4.xlsx";
        String columnRange = "E,A"; // Start column is greater than end column
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
           columnConverter.action(inputFilePath, columnRange, conversionType);
        });

        assertEquals("Start column cannot be greater than end column.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "test_files/non_existent.xlsx";
//        String outputFilePath = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Desktop\\ExcelColumnConvert\\src\\test\\resources\\test_files\\output5.xlsx";
        String columnRange = "A,B";
        String conversionType = "StringToNumber";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            columnConverter.action(inputFilePath, columnRange, conversionType);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testValidStringToNumberConversion2() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Column_Convert/output10.xlsx";
        String columnRange = "B,D";
        String conversionType = "StringToNumber";

        assertDoesNotThrow(() -> {
            String result = columnConverter.action(inputFilePath, columnRange, conversionType).get();
            assertTrue(result.contains(inputFilePath));
        });
    }
}

