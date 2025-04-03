package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ColorFormatTest {

    private ColorFormat colorFormat;

    @BeforeEach
    void setUp() {
        colorFormat = new ColorFormat();
    }

    @Test
    void testValidColoringByColumn() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output.xlsx";
        String condition = "Column";
        String range = "B";
        String colorIndex = "10"; // Example color index

        assertDoesNotThrow(() -> colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testValidColoringByRow() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output1.xlsx";
        String condition = "Row";
        String range = "5";
        String colorIndex = "12";

        assertDoesNotThrow(() -> colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testInvalidRangeFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output2.xlsx";
        String condition = "Column";
        String range = "1A"; // Invalid format
        String colorIndex = "10";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex);
        });
        assertTrue(exception.getMessage().contains("Invalid"));
    }

    @Test
    void testvalidColumnRangeFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output3.xlsx";
        String condition = "Column";
        String range = "A1:C5"; // Invalid format
        String colorIndex = "10";

        assertDoesNotThrow(() -> colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testvalidRowRangeFormat() {
        String inputFilePath = "src/test/resources/test_files/originals/valid_input.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output4.xlsx";
        String condition = "Row";
        String range = "1:5"; // Invalid format
        String colorIndex = "15";

        assertDoesNotThrow(() -> colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/non_existent.xlsx";
        String outputFilePath = "src/test/resources/test_files/ColorFormat/output5.xlsx";
        String condition = "Column";
        String range = "A";
        String colorIndex = "10";

        Exception exception = assertThrows(BotCommandException.class, () -> {
            colorFormat.action(inputFilePath, outputFilePath, condition, range, colorIndex);
        });
        assertTrue(exception.getMessage().contains("File processing error"));
    }

}
