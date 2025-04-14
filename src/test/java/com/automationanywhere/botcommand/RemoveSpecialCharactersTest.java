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

class RemoveSpecialCharactersTest {

    private RemoveSpecialCharacters removeSpecialCharacters;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/Remove_Special/";

    @BeforeEach
    void setUp() throws IOException {
        removeSpecialCharacters = new RemoveSpecialCharacters();
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
    void testValidSpecialCharacterRemoval() {
        String inputFilePath = workingDir+"special_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Special/output.xlsx";
        String column = "A";
        String specialCharacter = "-";

        assertDoesNotThrow(() -> {
            String result = removeSpecialCharacters.action(inputFilePath, column, specialCharacter).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());

    }

    @Test
    void testSpecialCharacterRemovalInDifferentColumn() {
        String inputFilePath = workingDir+"special_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Special/output1.xlsx";
        String column = "B";
        String specialCharacter = "#";

        assertDoesNotThrow(() -> {
            String result = removeSpecialCharacters.action(inputFilePath, column, specialCharacter).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testInvalidSpecialCharacterFormat() {
        String inputFilePath = workingDir+"special_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Special/output2.xlsx";
        String column = "A";
        String specialCharacter = "--"; // Invalid format, more than one character

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, column, specialCharacter);
        });

        assertEquals("Invalid input: Enter exactly one special character (e.g., -, @, /, #, etc.). Multiple characters are not allowed.", exception.getMessage());
    }

    @Test
    void testInvalidSpecialCharacterAsLetter() {
        String inputFilePath = workingDir+"special_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Special/output3.xlsx";
        String column = "A";
        String specialCharacter = "a"; // Invalid character, it cannot be a letter

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, column, specialCharacter);
        });

        assertEquals("Invalid input: Spacial character cannot be a letter.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = workingDir+"test_files/non_existent.xlsx";
//        String outputFilePath = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Desktop\\RemoveSpecialCharacters\\src\\test\\resources\\test_files\\output4.xlsx";
        String column = "A";
        String specialCharacter = "@";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, column, specialCharacter);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testValidSpecialCharacterRemovalWithDifferentColumn() {
        String inputFilePath = workingDir+"special_valid_input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Remove_Special/output5.xlsx";
        String column = "C";
        String specialCharacter = "$";

        assertDoesNotThrow(() -> {
            String result = removeSpecialCharacters.action(inputFilePath, column, specialCharacter).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }
}
