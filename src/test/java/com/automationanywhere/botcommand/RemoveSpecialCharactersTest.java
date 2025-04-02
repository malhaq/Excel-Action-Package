package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class RemoveSpecialCharactersTest {

    private RemoveSpecialCharacters removeSpecialCharacters;

    @BeforeEach
    void setUp() {
        removeSpecialCharacters = new RemoveSpecialCharacters();
    }

    @Test
    void testValidSpecialCharacterRemoval() {
        String inputFilePath = "src/test/resources/test_files/originals/special_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Special/output.xlsx";
        String column = "A";
        String specialCharacter = "-";

        assertDoesNotThrow(() -> removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testSpecialCharacterRemovalInDifferentColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/special_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Special/output1.xlsx";
        String column = "B";
        String specialCharacter = "#";

        assertDoesNotThrow(() -> removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testInvalidSpecialCharacterFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/special_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Special/output2.xlsx";
        String column = "A";
        String specialCharacter = "--"; // Invalid format, more than one character

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter);
        });

        assertEquals("Invalid input: Enter exactly one special character (e.g., -, @, /, #, etc.). Multiple characters are not allowed.", exception.getMessage());
    }

    @Test
    void testInvalidSpecialCharacterAsLetter() {
        String inputFilePath = "src/test/resources/test_files/originals/special_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Special/output3.xlsx";
        String column = "A";
        String specialCharacter = "a"; // Invalid character, it cannot be a letter

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter);
        });

        assertEquals("Invalid input: Spacial character cannot be a letter.", exception.getMessage());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "test_files/non_existent.xlsx";
        String outputFilePath = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Desktop\\RemoveSpecialCharacters\\src\\test\\resources\\test_files\\output4.xlsx";
        String column = "A";
        String specialCharacter = "@";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter);
        });

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testValidSpecialCharacterRemovalWithDifferentColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/special_valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Remove_Special/output5.xlsx";
        String column = "C";
        String specialCharacter = "$";

        assertDoesNotThrow(() -> removeSpecialCharacters.action(inputFilePath, outputFilePath, column, specialCharacter));
        assertTrue(new File(outputFilePath).exists());
    }
}
