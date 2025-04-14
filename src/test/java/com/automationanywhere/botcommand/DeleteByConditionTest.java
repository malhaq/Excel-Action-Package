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


public class DeleteByConditionTest {

    private DeleteByCondition deleteByCondition;
    private final String testResourcesDir = "src/test/resources/test_files/";
    private final String originalsDir = testResourcesDir + "original/";
    private final String workingDir = testResourcesDir + "working/Delete_By_Condition/";

    @BeforeEach
    void setUp() throws IOException {
        deleteByCondition = new DeleteByCondition();
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
    void testEqualsConditionWithStringColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/EqualsConditionWithStringColumnOutput.xlsx";
        String column = "A";
        String Condition = "Equals to =";
        String criteriaValue = "Fathi";//7 grace

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());


    }

    @Test
    void testDoesNotEqualsConditionWithStringColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotEqualsConditionWithStringColumn.xlsx";
        String column = "A";
        String Condition = "Doesn't equals !=";
        String criteriaValue = "Fathi";//only one left

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testEqualsConditionWithNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/EqualsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Equals to =";
        String criteriaValue = "19.0";//7 grace

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testDoesNotEqualsConditionWithNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotEqualsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Doesn't equals !=";
        String criteriaValue = "25";//only one left

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testContainsConditionWithStringColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ContainsConditionWithStringColumn.xlsx";
        String column = "D";
        String Condition = "Contain";
        String criteriaValue = "5th";

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testDoesNotContainConditionWithStringColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotContainConditionWithStringColumn.xlsx";
        String column = "D";
        String Condition = "Doesn't Contain";
        String criteriaValue = "th";

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testContainsConditionWithNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ContainsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Contain";
        String criteriaValue = "1";

//        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue));
//        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testDoesNotContainConditionWithNumericalColumn() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotContainConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Doesn't Contain";
        String criteriaValue = "9";

        assertDoesNotThrow(() -> {
            String result = deleteByCondition.action(inputFilePath, column, Condition, criteriaValue).get();
            assertTrue(result.contains(inputFilePath));
        });
        assertTrue(new File(inputFilePath).exists());
    }

    @Test
    void testFileNotFound() {
        String inputFilePath = "src/test/resources/test_files/original/NonExistentFile.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/FileNotFoundOutput.xlsx";
        String column = "A";
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnInput() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/InvalidColumnInputOutput.xlsx";
        String column = "1";  // Invalid column
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("Invalid column letter"));
    }

    @Test
    void testColumnInputIsSpace() {
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
//        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ColumnInputIsSpaceOutput.xlsx";
        String column = " ";  // Space instead of column letter
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("Invalid column letter"));
    }

}
