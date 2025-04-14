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

class RemoveDuplicatesTest {

    private RemoveDuplicates removeDuplicates;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/Remove_Duplicate/";

    @BeforeEach
    void setUp() throws IOException {
        removeDuplicates = new RemoveDuplicates();
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
    void testBasicFunctionalityStringColumn() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/basicFunctionalityStringColumnOutput.xlsx";
        String column = "E";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testBasicFunctionalityNumericalColumn() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/basicFunctionalityTestNumericalColumnOutput.xlsx";
        String column = "B";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testNoDuplicates() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/nonDuplicatesTestOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testColumnWithMixedDataTypes() {
        String inputFilePath = workingDir+"mixedDataTypes_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/mixedDataTypesOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testCaseSensitivityCheck() {
        String inputFilePath = workingDir+"caseSensitivity_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/caseSensitivityTestOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testLargeDataset() {
        String inputFilePath = workingDir+"largeDataset_input.xlsx";
        //test_files/original/largeDataset_input.xlsx
        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/largeDatasetTestOutput-V5.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result = removeDuplicates.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath =workingDir+"valid.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "A";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnLetter() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "#"; // Invalid character

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, column);
        });

        assertEquals("Invalid column letters: #", exception.getMessage());
    }

    @Test
    void testEmptyColumnLetter() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Duplicate/noOutput.xlsx";
        String column = "";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeDuplicates.action(inputFilePath, column);
        });

        assertEquals("Invalid column letters: ", exception.getMessage());
    }

}

