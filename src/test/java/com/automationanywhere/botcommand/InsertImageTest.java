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

class InsertImageTest {

    private InsertImage insertImage;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "originals/";
    private final String workingDir = testResourcesDir + "working/InsertImage/";

    @BeforeEach
    void setUp() throws IOException {
        insertImage = new InsertImage();

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
    void testInsertImageValid() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = originalsDir + "test_image.jpg";
        String targetCell = "B2";

        assertDoesNotThrow(() -> {
            String result = insertImage.action(excelFilePath, imagePath, targetCell).get();
            assertTrue(result.contains(excelFilePath));
        });

        assertTrue(new File(excelFilePath).exists());
    }

    @Test
    void testMissingExcelFile() {
        String excelFilePath = workingDir + "missing.xlsx";
        String imagePath = originalsDir + "test_image.jpg";
        String targetCell = "B2";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            insertImage.action(excelFilePath, imagePath, targetCell);
        });

        assertTrue(exception.getMessage().contains("Excel file does not exist"));
    }

    @Test
    void testMissingImageFile() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = workingDir + "missing_image.png";
        String targetCell = "A1";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            insertImage.action(excelFilePath, imagePath, targetCell);
        });

        assertTrue(exception.getMessage().contains("Image file does not exist"));
    }

    @Test
    void testInvalidImage() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = originalsDir + "invalid_image.jpg";
        String targetCell = "C3";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            insertImage.action(excelFilePath, imagePath, targetCell);
        });

        assertTrue(exception.getMessage().contains("Image file does not exist:"));
    }

    @Test
    void testInvalidImageFormat() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = originalsDir + "test_image.bmp";
        String targetCell = "C3";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            insertImage.action(excelFilePath, imagePath, targetCell);
        });
        System.out.println(exception.getMessage());
        assertTrue(exception.getMessage().contains("Only PNG and JPG images are supported"));
    }

    @Test
    void testInvalidCellFormat() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = originalsDir + "test_image.jpg";
        String targetCell = "3B"; // Invalid cell

        Exception exception = assertThrows(BotCommandException.class, () -> {
            insertImage.action(excelFilePath, imagePath, targetCell);
        });

        assertTrue(exception.getMessage().contains("Target cell format is invalid"));
    }

    @Test
    void testInsertImageInRange() {
        String excelFilePath = workingDir + "valid_input 1.xlsx";
        String imagePath = originalsDir + "test_image.jpg";
        String targetCell = "A10:B15";

        assertDoesNotThrow(() -> {
            String result = insertImage.action(excelFilePath, imagePath, targetCell).get();
            assertTrue(result.contains(excelFilePath));
        });

        assertTrue(new File(excelFilePath).exists());
    }

    @Test
    void testInsertImageInCell() {
        String excelFilePath = workingDir + "valid_input.xlsx";
        String imagePath = originalsDir + "test_image.jpg";
        String targetCell = "AD10";

        assertDoesNotThrow(() -> {
            String result = insertImage.action(excelFilePath, imagePath, targetCell).get();
            assertTrue(result.contains(excelFilePath));
        });

        assertTrue(new File(excelFilePath).exists());
    }

}
