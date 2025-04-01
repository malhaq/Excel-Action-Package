package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;
import static org.junit.jupiter.api.Assertions.*;

class DeleteBlankTest {

    private DeleteBlank deleteBlank;

    @BeforeEach
    void setUp() {
        deleteBlank = new DeleteBlank();
    }

    @Test
    void testBasicFunctionalityStringColumn() {
        String inputFilePath = "src/test/resources/test_files/original/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/basicFunctionalityStringColumnOutput.xlsx";
        String column = "E";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testBasicFunctionalityNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/original/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/basicFunctionalityNumericalColumnOutput.xlsx";
        String column = "B";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testNoBlankCells() {
        String inputFilePath = "src/test/resources/test_files/original/no_blanks_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noBlanksOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testAllBlankCells() {
        String inputFilePath = "src/test/resources/test_files/original/all_blanks_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/allBlanksOutput.xlsx";
        String column = "C";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testMixedBlankCells() {
        String inputFilePath = "src/test/resources/test_files/original/mixed_blanks_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/mixedBlanksOutput.xlsx";
        String column = "D";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testLargeDataset() {
        String inputFilePath = "src/test/resources/test_files/original/large_dataset_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/largeDatasetOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/nonexistent_file.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noOutput.xlsx";
        String column = "A";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, outputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnLetter() {
        String inputFilePath = "src/test/resources/test_files/original/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/oOutput.xlsx";
        String column = "#";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, outputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("Invalid column letters: "));
    }

    @Test
    void testEmptyColumnLetter() {
        String inputFilePath = "src/test/resources/test_files/original/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noOutput.xlsx";
        String column = "";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, outputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("Invalid column letters: "));
    }

    @Test
    void testEmptyFile() {
        String inputFilePath = "src/test/resources/test_files/original/empty_file.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_Blank/emptyFileOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> deleteBlank.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }
}