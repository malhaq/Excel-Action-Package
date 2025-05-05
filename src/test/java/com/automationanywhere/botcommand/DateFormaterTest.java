package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.*;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class DateFormaterTest {

    private DateFormater dateFormater;

    @BeforeEach
    void setUp() throws IOException {
        dateFormater = new DateFormater();

    }



    @Test
    void testActionSuccess() {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\Date_Formater\\test_dates.xlsx";
        Value<String> result = dateFormater.action(
                testFile,
                "A",
                "yyyy-MM-dd",
                true
        );

        assertNotNull(result);
        assertEquals("Date formatting completed successfully", result.get());
    }

    @Test
    void testActionWithInvalidFilePath() {
        assertThrows(BotCommandException.class, () -> {
            dateFormater.action(
                    "nonexistent_file.xlsx",
                    "A",
                    "MMMM d, yyyy",
                    false
            );
        });
    }

    @Test
    void testActionWithInvalidColumn() {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\Date_Formater\\test_dates.xlsx";
        assertThrows(BotCommandException.class, () -> {
            dateFormater.action(
                    testFile,
                    "ZZ",
                    "MM-dd-yyyy",
                    true
            );
        });
    }

    @Test
    void testActionWithEmptySheet() throws IOException {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\Date_Formater\\test_dates2.xlsx";
        assertThrows(BotCommandException.class, () -> {
            dateFormater.action(
                    testFile,
                    "A",
                    "MM-dd-yyyy",
                    true
            );
        });
    }

    @Test
    void testFormatColumnWithHeaders() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\Date_Formater\\test_dates1.xlsx";
        Value<String> result = dateFormater.action(
                testFile,
                "A",
                "yyyy-MM-dd",
                true
        );
        assertNotNull(result);
        assertEquals("Date formatting completed successfully", result.get());

    }

    @Test
    void testFormatColumnWithoutHeaders() throws Exception {
        String testFile = "C:\\Users\\malhaq\\OneDrive - Air Arabia Group\\Documents\\ExcelColumnConvert\\ExcelColumnConvert\\src\\test\\resources\\test_files\\Date_Formater\\test_dates_no_headers.xlsx";

        Value<String> result = dateFormater.action(
                testFile,
                "A",
                "EEEE, MMM d",
                false
        );

        assertNotNull(result);
        assertEquals("Date formatting completed successfully", result.get());

    }


}