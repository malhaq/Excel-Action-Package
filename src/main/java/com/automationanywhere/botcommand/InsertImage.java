package com.automationanywhere.botcommand;


import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;



import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "InserteImage", label = "Insert image in a cell", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "Updated File Path", return_type = STRING, return_required = true)


public class InsertImage {
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(description = "Provide the full path of the Excel file that contains the data to be converted.", node_label = "Select the Excel file to process", label = "Excel File Path")
            @NotEmpty
            String inputFilePath,

            @Idx(index = "2", type = FILE)
            @Pkg(
                    label = "Image File Path",
                    description = "Provide the full path of the image to process.",
                    node_label = "Select the image file"
            )
            @NotEmpty
            String imagePath,

            @Idx(index = "3", type = TEXT)
            @Pkg(
                    label = "Target Cell Range",//A5,B11,AB6
                    description = "Enter the cell location to insert the image, e.g., A1, AB6, A1:D10.",
                    node_label = "Cell Range to insert the image into"
            )
            @NotEmpty
            String cell
    ) {

        // Validation
        if (inputFilePath == null || inputFilePath.trim().isEmpty())
            throw new BotCommandException("Excel file path cannot be empty");

        if (!new File(inputFilePath).exists())
            throw new BotCommandException("Excel file does not exist: " + inputFilePath);

        if (imagePath == null || imagePath.trim().isEmpty())
            throw new BotCommandException("Image file path cannot be empty");

        if (!new File(imagePath).exists())
            throw new BotCommandException("Image file does not exist: " + imagePath);

        if (!(imagePath.toLowerCase().endsWith(".png") || imagePath.toLowerCase().endsWith(".jpg") || imagePath.toLowerCase().endsWith(".jpeg")))
            throw new BotCommandException("Only PNG and JPG images are supported");

        File excelFile = new File(inputFilePath);
        if (!excelFile.exists() || !excelFile.isFile()) {
            throw new BotCommandException("Excel file does not exist or is not a valid file: " + inputFilePath);
        }
        if (!inputFilePath.toLowerCase().endsWith(".xlsx")) {
            throw new BotCommandException("Only .xlsx files are supported for Excel input.");
        }

        File imgFile = new File(imagePath);
        if (!imgFile.exists() || !imgFile.isFile()) {
            throw new BotCommandException("Image file does not exist or is not a valid file: " + imagePath);
        }
        if (!imagePath.toLowerCase().matches(".*\\.(jpg|jpeg|png)$")) {
            throw new BotCommandException("Only JPG, JPEG, and PNG image formats are supported.");
        }

        if (cell == null || cell.trim().isEmpty()) {
            throw new BotCommandException("Target cell cannot be empty.");
        }

        int startColumn, startRow, endColumn, endRow;

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             FileInputStream imageInputStream = new FileInputStream(imagePath)) {

            if (!imagePath.toLowerCase().matches(".*\\.(jpg|jpeg|png|bmp|gif)$")) {
                throw new RuntimeException("Unsupported file type. Please upload a valid image.");
            }
            if (validateCellOrRange(cell)) {
                if(cell.contains(":")){
                    String [] array = cell.split(":");

                    int digitIndx = ExcelUtils.firstDigitIndex(array[0]);
                    startColumn = ExcelUtils.columnLetterToIndex(array[0].substring(0, digitIndx));
                    startRow = Integer.parseInt(array[0].substring(digitIndx).trim()) - 1;

                    digitIndx = ExcelUtils.firstDigitIndex(array[1]);
                    endColumn = ExcelUtils.columnLetterToIndex(array[1].substring(0, digitIndx));
                    endRow = Integer.parseInt(array[1].substring(digitIndx).trim()) - 1;
                }else {
                    int digitIndx = ExcelUtils.firstDigitIndex(cell);
                    startColumn = ExcelUtils.columnLetterToIndex(cell.substring(0, digitIndx));
                    startRow = Integer.parseInt(cell.substring(digitIndx).trim()) - 1;
                    endColumn = startColumn+1;
                    endRow = startRow+1;
                }
            } else {
                String [] array = cell.split(":");
                throw new RuntimeException("Cell formate is wrong");
            }


            XSSFSheet sheet = workbook.getSheetAt(0);

            // Read image into byte array
            byte[] bytes = IOUtils.toByteArray(imageInputStream);
            int pictureType = Workbook.PICTURE_TYPE_PNG;
            if (imagePath.toLowerCase().endsWith(".jpg") || imagePath.toLowerCase().endsWith(".jpeg")) {
                pictureType = Workbook.PICTURE_TYPE_JPEG;
            }
            int pictureIdx = workbook.addPicture(bytes, pictureType);

            // Create drawing and anchor
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor();
            anchor.setCol1(startColumn); // Column start
            anchor.setRow1(startRow); // Row start
            anchor.setCol2(endColumn + 1); // Column end (optional: size)
            anchor.setRow2(endRow + 1); // Row end (optional: size)

            drawing.createPicture(anchor, pictureIdx);

            try (FileOutputStream fos = new FileOutputStream(inputFilePath)) {
                workbook.write(fos);
            }

            return new StringValue("Conversion completed. Saved as: " + inputFilePath);
        } catch (Exception e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }
    private boolean validateCellOrRange(String input) {
        String singleCellPattern = "^[A-Z]+[1-9][0-9]*$";
        String rangePattern = "^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$";

        if (!input.matches(singleCellPattern) && !input.matches(rangePattern)) {
            return false;
        }
        return true;
    }

}
