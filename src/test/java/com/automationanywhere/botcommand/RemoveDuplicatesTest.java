package com.automationanywhere.botcommand;



import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;
import static org.junit.jupiter.api.Assertions.*;

class RemoveDuplicatesTest {

    private RemoveDuplicates removeDuplicates;

    @BeforeEach
    void setUp() {
        removeDuplicates = new RemoveDuplicates();
    }

    @Test
    void testBasicFunctionalityStringColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/basicFunctionalityStringColumnOutput.xlsx";
        String column = "E";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testBasicFunctionalityNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/basicFunctionalityTestNumericalColumnOutput.xlsx";
        String column = "B";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testNoDuplicates() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/nonDuplicatesTestOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testColumnWithMixedDataTypes() {
        String inputFilePath = "src/test/resources/test_files/originals/mixedDataTypes_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/mixedDataTypesOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testCaseSensitivityCheck() {
        String inputFilePath = "src/test/resources/test_files/originals/caseSensitivity_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/caseSensitivityTestOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testLargeDataset() {
        String inputFilePath = "src/test/resources/test_files/originals/largeDataset_input.xlsx";
        //test_files/original/largeDataset_input.xlsx
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/largeDatasetTestOutput-V5.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> removeDuplicates.action(inputFilePath, outputFilePath, column));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Desktop\\ExcelColumnConvert\\src\\test\\resources\\test_files\\valid.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "A";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, outputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnLetter() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "#"; // Invalid character

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, outputFilePath, column);
        });

        assertEquals("Invalid column letters: #", exception.getMessage());
    }

    @Test
    void testEmptyColumnLetter() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, outputFilePath, column);
        });

        assertEquals("Invalid column letters: ", exception.getMessage());
    }

}

