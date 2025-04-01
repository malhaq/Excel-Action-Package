package com.automationanywhere.botcommand;

import com.automationanywhere.botcommand.exception.BotCommandException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.File;
import static org.junit.jupiter.api.Assertions.*;


public class DeleteByConditionTest {

    private DeleteByCondition deleteByCondition;

    @BeforeEach
    void setUp() {
        deleteByCondition = new DeleteByCondition();
    }

    @Test
    void testEqualsConditionWithStringColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/EqualsConditionWithStringColumnOutput.xlsx";
        String column = "A";
        String Condition = "Equals to =";
        String criteriaValue = "Fathi";//7 grace

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column,Condition,criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testDoesNotEqualsConditionWithStringColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotEqualsConditionWithStringColumn.xlsx";
        String column = "A";
        String Condition = "Doesn't equals !=";
        String criteriaValue = "Fathi";//only one left

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column,Condition,criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testEqualsConditionWithNumericalColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/EqualsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Equals to =";
        String criteriaValue = "19.0";//7 grace

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column,Condition,criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testDoesNotEqualsConditionWithNumericalColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotEqualsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Doesn't equals !=";
        String criteriaValue = "25";//only one left

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column,Condition,criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testContainsConditionWithStringColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ContainsConditionWithStringColumn.xlsx";
        String column = "D";
        String Condition = "Contain";
        String criteriaValue = "5th";

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testDoesNotContainConditionWithStringColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotContainConditionWithStringColumn.xlsx";
        String column = "D";
        String Condition = "Doesn't Contain";
        String criteriaValue = "th";

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testContainsConditionWithNumericalColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ContainsConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Contain";
        String criteriaValue = "1";

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testDoesNotContainConditionWithNumericalColumn(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/DoesNotContainConditionWithNumericalColumn.xlsx";
        String column = "B";
        String Condition = "Doesn't Contain";
        String criteriaValue = "9";

        assertDoesNotThrow(() -> deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue));
        assertTrue(new File(outputFilePath).exists());
    }

    @Test
    void testFileNotFound(){
        String inputFilePath = "src/test/resources/test_files/original/NonExistentFile.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/FileNotFoundOutput.xlsx";
        String column = "A";
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("File processing error"));
    }

    @Test
    void testInvalidColumnInput(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/InvalidColumnInputOutput.xlsx";
        String column = "1";  // Invalid column
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("Invalid column letter"));
    }

    @Test
    void testColumnInputIsSpace(){
        String inputFilePath = "src/test/resources/test_files/original/Test_Input.xlsx";
        String outputFilePath = "src/test/resources/test_files/Delete_By_Condition/ColumnInputIsSpaceOutput.xlsx";
        String column = " ";  // Space instead of column letter
        String Condition = "Equals to =";
        String criteriaValue = "fathi";

        Exception exception = assertThrows(BotCommandException.class, () ->
                deleteByCondition.action(inputFilePath, outputFilePath, column, Condition, criteriaValue)
        );
        assertTrue(exception.getMessage().contains("Invalid column letter"));
    }

}
