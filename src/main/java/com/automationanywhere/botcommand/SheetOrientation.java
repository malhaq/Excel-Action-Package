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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be dispalable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "SheetOrientation", label = "Change Sheet Orientation", node_label = "Change Orientation of a Sheet", description = "Converts the orientation of the first sheet in an Excel file to landscape or portrait and sets the page size.", icon = "excel_icon.svg",

        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_label = "Updated File Path", return_type = STRING, return_required = true)

public class SheetOrientation {
    private static final Logger LOGGER = Logger.getLogger(ColumnConverter.class.getName());

    //Identify the entry point for the action. Returns a Value<String> because the return type is String.
    @Execute
    public Value<String> action(
            @Idx(index = "1", type = FILE)
            @Pkg(description = "Provide the full path of the Excel file that contains the data to be converted.",node_label = "Select the Excel file to process",label = "Excel File Path")
            @NotEmpty
            String inputFilePath,

            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1",pkg = @Pkg(label = "Convert to Land Scape", value = "ToLandScape")),
                    @Idx.Option(index = "2.2",pkg = @Pkg(label = "Convert to Land Portrait", value = "ToPortrait"))
            })
            @Pkg(label = "Sheet Orientation")
            @NotEmpty
            String conversionType,

            @Idx(index = "3", type = SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "PRINTER_DEFAULT", value = "0")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "LETTER", value = "1")),
                    @Idx.Option(index = "3.3", pkg = @Pkg(label = "LETTER_SMALL", value = "2")),
                    @Idx.Option(index = "3.4", pkg = @Pkg(label = "TABLOID", value = "3")),
                    @Idx.Option(index = "3.5", pkg = @Pkg(label = "LEDGER", value = "4")),
                    @Idx.Option(index = "3.6", pkg = @Pkg(label = "LEGAL", value = "5")),
                    @Idx.Option(index = "3.7", pkg = @Pkg(label = "STATEMENT", value = "6")),
                    @Idx.Option(index = "3.8", pkg = @Pkg(label = "EXECUTIVE", value = "7")),
                    @Idx.Option(index = "3.9", pkg = @Pkg(label = "A3", value = "8")),
                    @Idx.Option(index = "3.10", pkg = @Pkg(label = "A4", value = "9")),
                    @Idx.Option(index = "3.11", pkg = @Pkg(label = "A4_SMALL", value = "10")),
                    @Idx.Option(index = "3.12", pkg = @Pkg(label = "A5", value = "11")),
                    @Idx.Option(index = "3.13", pkg = @Pkg(label = "B4", value = "12")),
                    @Idx.Option(index = "3.14", pkg = @Pkg(label = "B5", value = "13")),
                    @Idx.Option(index = "3.15", pkg = @Pkg(label = "FOLIO8", value = "14")),
                    @Idx.Option(index = "3.16", pkg = @Pkg(label = "QUARTO", value = "15")),
                    @Idx.Option(index = "3.17", pkg = @Pkg(label = "TEN_BY_FOURTEEN", value = "16")),
                    @Idx.Option(index = "3.18", pkg = @Pkg(label = "ELEVEN_BY_SEVENTEEN", value = "17")),
                    @Idx.Option(index = "3.19", pkg = @Pkg(label = "NOTE8", value = "18")),
                    @Idx.Option(index = "3.20", pkg = @Pkg(label = "ENVELOPE_9", value = "19")),
                    @Idx.Option(index = "3.21", pkg = @Pkg(label = "ENVELOPE_10", value = "20")),
                    @Idx.Option(index = "3.22", pkg = @Pkg(label = "ENVELOPE_DL", value = "27")),
                    @Idx.Option(index = "3.23", pkg = @Pkg(label = "ENVELOPE_CS", value = "28")),
                    @Idx.Option(index = "3.24", pkg = @Pkg(label = "ENVELOPE_C5", value = "28")),
                    @Idx.Option(index = "3.25", pkg = @Pkg(label = "ENVELOPE_C3", value = "29")),
                    @Idx.Option(index = "3.26", pkg = @Pkg(label = "ENVELOPE_C4", value = "30")),
                    @Idx.Option(index = "3.27", pkg = @Pkg(label = "ENVELOPE_C6", value = "31")),
                    @Idx.Option(index = "3.28", pkg = @Pkg(label = "ENVELOPE_MONARCH", value = "37")),
                    @Idx.Option(index = "3.29", pkg = @Pkg(label = "A4_EXTRA", value = "53")),
                    @Idx.Option(index = "3.30", pkg = @Pkg(label = "A4_TRANSVERSE", value = "55")),
                    @Idx.Option(index = "3.31", pkg = @Pkg(label = "A4_PLUS", value = "60")),
                    @Idx.Option(index = "3.32", pkg = @Pkg(label = "LETTER_ROTATED", value = "75")),
                    @Idx.Option(index = "3.33", pkg = @Pkg(label = "A4_ROTATED", value = "77"))


            })
            @Pkg(label = "Page Size")
            @NotEmpty
            String pageSize
    ) {
        short paperSize;
        try (Workbook workbook = ExcelUtils.openWorkbook(inputFilePath)) {

            Sheet sheet = workbook.getSheetAt(0);
            PrintSetup printSetup = sheet.getPrintSetup();
            if(conversionType.equals("LandScape")) {
                printSetup.setLandscape(true);
            } else if (conversionType.equals("Portrait")) {
                printSetup.setLandscape(false);
            }else {
                throw new RuntimeException("Unknown conversion type");
            }
            paperSize = Short.parseShort(pageSize);
            printSetup.setPaperSize(paperSize);
            try (FileOutputStream fos = new FileOutputStream(inputFilePath)) {
                workbook.write(fos);
                fos.close();
                workbook.close();
            }

            return new StringValue("Conversion completed. Saved as: " + inputFilePath);
        } catch (Exception e) {
            throw new BotCommandException("File processing error: " + e.getMessage());
        }
    }

}
