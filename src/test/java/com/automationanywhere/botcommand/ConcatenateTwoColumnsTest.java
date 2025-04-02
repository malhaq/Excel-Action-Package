package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ConcatenateTwoColumnsTest {

    private ConcatenateTwoColumns concatenateTwoColumns;

    @BeforeEach
    void setUp() {
        concatenateTwoColumns = new ConcatenateTwoColumns();
    }

    @Test
    void testValidConcatenation() {
        String inputFilePath = "src/test/resources/test_files/originals/Concatenate_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testInvalidColumnFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/Concatenate_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output1.xlsx";
        String firstColumn = "1"; // Invalid format
        String secondColumn = "B";
        String outputColumn = "C";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertEquals("Invalid column letters: 1", exception.getMessage());
    }

    @Test
    void testOutputColumnSameAsInputColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/Concatenate_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output2.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "A"; // Output column same as input column

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertEquals("Output column cannot be the same as an input column.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/non_existent.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output3.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn);
        });
        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testEmptyCells() {
        String inputFilePath = "src/test/resources/test_files/originals/empty_file.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output4.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testNumbersConcatenation() {
        String inputFilePath = "src/test/resources/test_files/originals/numeric_values.xlsx";
        String outputFilePath = "src/test/resources/test_files/Concatenation/output5.xlsx";
        String firstColumn = "A";
        String secondColumn = "B";
        String outputColumn = "C";

        assertDoesNotThrow(() -> concatenateTwoColumns.action(inputFilePath, outputFilePath, firstColumn, secondColumn, outputColumn));
        assertTrue(new File(outputFilePath).exists());
    }
}
