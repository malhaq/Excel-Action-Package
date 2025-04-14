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

class ConcatenateTwoColumnsTest {
    private ConcatenateTwoColumns concatenateTwoColumns;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/Concatenation/";

    @BeforeEach
    void setUp() throws IOException {
        concatenateTwoColumns = new ConcatenateTwoColumns();

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
    void testValidConcatenation() {
        String inputFilePath = workingDir+"Concatenate_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Concatenation/output.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> {
            String result = concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testInvalidColumnFormat() {
        String inputFilePath = workingDir + "Concatenate_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output1.xlsx";
        String firstColumn = "1"; // Invalid format
        String secondColumn = "B";
        String outputColumn = "C";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertEquals("Invalid column letters: 1", exception.getMessage());
    }

    @Test
    void testOutputColumnSameAsInputColumn() {
        String inputFilePath = workingDir + "Concatenate_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Concatenation/output2.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "A"; // Output column same as input column

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertEquals("Output column cannot be the same as an input column.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = workingDir + "non_existent.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Concatenation/output3.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testEmptyCells() {
        String inputFilePath = workingDir + "empty_file.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Concatenation/output4.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> {
            String result = concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testNumbersConcatenation() {
        String inputFilePath = workingDir + "numeric_values.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Concatenation/output5.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> {
            String result = concatenateTwoColumns.action(inputFilePath, firstColumn, secondColumn, outputColumn).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }
}
