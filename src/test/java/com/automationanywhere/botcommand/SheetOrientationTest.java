package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import static org.junit.jupiter.api.Assertions.*;

class SheetOrientationTest {

    private SheetOrientation sheetOrientation;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/SheetOrientation/";
    private final String testFileName = "valid_input.xlsx";

    @BeforeEach
    void setUp() throws IOException {
        sheetOrientation = new SheetOrientation();

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
    void testConvertToLandscape() {
        String inputFilePath = workingDir + "valid_input 3.xlsx";
        String orientation = "LandScape";
        String paperSize = "1"; // LETTER

        assertDoesNotThrow(() -> {
            Value<String> result = sheetOrientation.action(inputFilePath, orientation, paperSize);
            assertNotNull(result);
            assertTrue(result.get().contains("Conversion completed"));
        });

        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testConvertToPortrait() {
        String inputFilePath = workingDir + "valid_input.xlsx";
        String orientation = "Portrait";
        String paperSize = "9"; // A4

        assertDoesNotThrow(() -> {
            Value<String> result = sheetOrientation.action(inputFilePath, orientation, paperSize);
            assertNotNull(result);
            assertTrue(result.get().contains("Conversion completed"));
        });

        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testInvalidOrientationThrowsException() {
        String inputFilePath = workingDir + testFileName;
        String invalidOrientation = "InvalidType";
        String paperSize = "1";

        BotCommandException exception = assertThrows(
                BotCommandException.class,
                () -> sheetOrientation.action(inputFilePath, invalidOrientation, paperSize)
        );

        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testFileNotFoundThrowsException() {
        String nonExistentFilePath = "src/test/resources/test_files/non_existent.xlsx";
        String orientation = "LandScape";
        String paperSize = "1";

        BotCommandException exception = assertThrows(
                BotCommandException.class,
                () -> sheetOrientation.action(nonExistentFilePath, orientation, paperSize)
        );

        assertTrue(exception.getMessage().contains("File processing error"));
    }
}
