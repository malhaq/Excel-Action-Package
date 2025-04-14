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

class DeleteBlankTest {

    private DeleteBlank deleteBlank;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "original/";
    private final String workingDir = testResourcesDir + "working/Delete_Blank/";

    @BeforeEach
    void setUp() throws IOException {
        deleteBlank = new DeleteBlank();

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
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/basicFunctionalityStringColumnOutput.xlsx";
        String column = "E";

        assertDoesNotThrow(() -> {
            String result = deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());

    }

    @Test
    void testBasicFunctionalityNumericalColumn() {
        String inputFilePath =workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/basicFunctionalityNumericalColumnOutput.xlsx";
        String column = "B";

        assertDoesNotThrow(() -> {
            String result = deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testNoBlankCells() {
        String inputFilePath = workingDir+"no_blanks_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noBlanksOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result =  deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testAllBlankCells() {
        String inputFilePath = workingDir+"all_blanks_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/allBlanksOutput.xlsx";
        String column = "C";

        assertDoesNotThrow(() -> {
            String result =  deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testMixedBlankCells() {
        String inputFilePath = workingDir+"mixed_blanks_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/mixedBlanksOutput.xlsx";
        String column = "D";

        assertDoesNotThrow(() -> {
            String result =  deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testLargeDataset() {
        String inputFilePath = workingDir+"large_dataset_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/largeDatasetOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result =  deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = workingDir+"nonexistent_file.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noOutput.xlsx";
        String column = "A";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnLetter() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/oOutput.xlsx";
        String column = "#";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("Invalid column letters: "));
    }

    @Test
    void testEmptyColumnLetter() {
        String inputFilePath = workingDir+"valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/noOutput.xlsx";
        String column = "";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            deleteBlank.action(inputFilePath, column);
        });

        assertTrue(exception.getMessage().contains("Invalid column letters: "));
    }

    @Test
    void testEmptyFile() {
        String inputFilePath = workingDir+"empty_file.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_Blank/emptyFileOutput.xlsx";
        String column = "A";

        assertDoesNotThrow(() -> {
            String result =  deleteBlank.action(inputFilePath, column).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }
}