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
    private static final Logger LOGGER = Logger.getLogger(DeleteByCondition.class.getName());

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Input File",description = "Path to the Excel or CSV file to be converted.")
            @NotEmpty
            String inputFilePath,

            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Excel (.xls) to Excel (.xlsx)", value = "XLStoXLSX")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Excel (.xlsx) to CSV", value = "XLSXtoCSV")),
                    @Idx.Option(index = "2.3", pkg = @Pkg(label = "Excel (.xls) to CSV",  value = "XLStoCSV")),
                    @Idx.Option(index = "2.4", pkg = @Pkg(label = "CSV to Excel (.xlsx)", value = "CSVtoXLSX"))

            })
            @Pkg(label = "Conversion Type", description = "Select the type of conversion to perform.")
            @NotEmpty
            String conversionType



    ) {

        try {
            if(!Files.exists(Paths.get(inputFilePath))) {
                throw new IOException("File not found: ");
            }

            switch (conversionType) {
                case "XLSXtoCSV":
                    ExcelUtils.convertXlsxToCsv(inputFilePath);
                    break;
                case "XLStoCSV":
                    ExcelUtils.convertXlsToCsv(inputFilePath);
                    break;
                case "XLStoXLSX":
                    ExcelUtils.convertXlsToXlsx(inputFilePath);
                    break;
                case "CSVtoXLSX":
                    ExcelUtils.convertCsvToXlsx(inputFilePath);
                    break;

            }
            return new StringValue("Original file successfully converted and overwritten.");
        } catch (IOException e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }

}
