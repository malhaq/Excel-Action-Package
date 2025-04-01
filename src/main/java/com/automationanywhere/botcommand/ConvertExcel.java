package com.automationanywhere.botcommand;


import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.ss.formula.atp.Switch;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "ConvertExcelAndCSV", label = "[[ConvertExcel.label]]", node_label = "[[ConvertExcel.node_label]]", description = "[[ConvertExcel.description]]", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "[[ConvertExcel.return_label]]", return_type = STRING, return_required = true)


public class ConvertExcel {
    private static final Messages MESSAGES = MessagesFactory.getMessages("com.automationanywhere.botcommand.samples.messages");
    private static final Logger LOGGER = Logger.getLogger(DeleteByCondition.class.getName());

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input File Path")
            @NotEmpty
            String inputFilePath,


            @Idx(index = "2", type = FILE)
            @Pkg(label = "Output File Path")
            @NotEmpty
            String outputFilePath,

            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "XLS to XLSX", node_label = "", value = "XLStoXLSX")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "XLSX to CSV", node_label = "", value = "XLSXtoCSV")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "XLS to CSV", node_label = "", value = "XLStoCSV")),
                    @Idx.Option(index = "3.4", pkg = @Pkg(label = "CSV to XLSX", node_label = "", value = "CSVtoXLSX"))
            })
            @Pkg(label = "Conversion Type")
            @NotEmpty
            String conversionType



    ) {

        try {
            if(!Files.exists(Paths.get(inputFilePath))) {
                throw new IOException("File not found: ");
            }

            switch (conversionType) {
                case "XLSXtoCSV":
                    ExcelUtils.convertXlsxToCsv(inputFilePath, outputFilePath);
                    break;
                case "XLStoCSV":
                    ExcelUtils.convertXlsToCsv(inputFilePath, outputFilePath);
                    break;
                case "XLStoXLSX":
                    ExcelUtils.convertXlsToXlsx(inputFilePath, outputFilePath);
                    break;
                case "CSVtoXLSX":
                    ExcelUtils.convertCsvToXlsx(inputFilePath, outputFilePath);
                    break;

            }
            return new StringValue("Output Saved as: " + outputFilePath);
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }

}
