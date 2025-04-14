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

class ColorFormatTest {

    private ColorFormat colorFormat;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "ColorFormat/";
    private final String workingDir = testResourcesDir + "working/";

    @BeforeEach
    void setUp() throws IOException {
        colorFormat = new ColorFormat();

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
    void testValidColoringByColumn() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Column";
        String range = "B";
        String colorIndex = "10"; // RED

        assertDoesNotThrow(() -> {
            String result = colorFormat.action(inputFilePath, condition, range, colorIndex).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testValidColoringByRow() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Row";
        String range = "5";
        String colorIndex = "12"; // BLUE

        assertDoesNotThrow(() -> {
            String result = colorFormat.action(inputFilePath, condition, range, colorIndex).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testInvalidRangeFormat() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Column";
        String range = "1A"; // Invalid format
        String colorIndex = "10"; // RED

        Exception exception = assertThrows(BotCommandException.class, () -> {
            colorFormat.action(inputFilePath, condition, range, colorIndex);
        });
        assertTrue(exception.getMessage().contains("Row value must be a number"));
    }

    @Test
    void testvalidColumnRangeFormat() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Column";
        String range = "A1:C5";
        String colorIndex = "10"; // RED

        assertDoesNotThrow(() -> {
            String result = colorFormat.action(inputFilePath, condition, range, colorIndex).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testvalidRowRangeFormat() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Row";
        String range = "1:5";
        String colorIndex = "15"; // TURQUOISE

        assertDoesNotThrow(() -> {
            String result = colorFormat.action(inputFilePath, condition, range, colorIndex).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/non_existent.xlsx";
//        String outputFilePath = "src/test/resources/test_files/ColorFormat/output5.xlsx";
        String condition = "Column";
        String range = "A";
        String colorIndex = "10";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            colorFormat.action(inputFilePath, condition, range, colorIndex);
        });
        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testSingleCellColoring() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String condition = "Column";
        String range = "B2";
        String colorIndex = "13"; // YELLOW

        assertDoesNotThrow(() -> {
            String result = colorFormat.action(inputFilePath, condition, range, colorIndex).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }
}


