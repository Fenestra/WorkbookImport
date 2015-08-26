package com.westat;

/**
 * Created by lee on 8/25/2015.
 */
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToJson {

    private Workbook workbook = null;
    private ArrayList<ArrayList<String>> JsonData = null;
    private int maxRowWidth = 0;
    private int formattingConvention = 0;
    private DataFormatter formatter = null;
    private FormulaEvaluator evaluator = null;
    private String separator = null;
    private String sheetName = null;
    private static final String Json_FILE_EXTENSION = ".Json";
    private StringBuffer sheets = new StringBuffer();
    private static final String DEFAULT_SEPARATOR = ",";
    public static final int EXCEL_STYLE_ESCAPING = 0;

    public void convertExcelToJson(String strSource, String strDestination)
            throws FileNotFoundException, IOException,
            IllegalArgumentException, InvalidFormatException {

        separator = DEFAULT_SEPARATOR;
        formattingConvention = EXCEL_STYLE_ESCAPING;

        File source = new File(strSource);
        File destination = new File(strDestination);
        String info = "convertExceltoJson source is " + source.getAbsolutePath() + " dest is " + destination.getAbsolutePath();
        File[] filesList = null;
        String destinationFilename = null;
System.out.println(info);
        // Check that the source file/folder exists.
        if (!source.exists()) {
            throw new IllegalArgumentException("The source for the Excel "
                    + "file(s) cannot be found." + info);
        }

        if (!destination.exists()) {
            throw new IllegalArgumentException("The folder/directory for the "
                    + "converted Json file(s) does not exist." + info);
        }
        if (!destination.isDirectory()) {
            throw new IllegalArgumentException("The destination for the Json "
                    + "file(s) is not a directory/folder.");
        }

        openWorkbook(source);

        convertToJson();

        destinationFilename = source.getName();
        destinationFilename = destinationFilename.substring(
                0, destinationFilename.lastIndexOf("."))
                + Json_FILE_EXTENSION;

        saveJsonFile(new File(destination, destinationFilename));

    }

    /**
     * Open an Excel workbook ready for conversion.
     *
     * @param file An instance of the File class that encapsulates a handle
     *        to a valid Excel workbook. Note that the workbook can be in
     *        either binary (.xls) or SpreadsheetML (.xlsx) format.
     * @throws java.io.FileNotFoundException Thrown if the file cannot be located.
     * @throws java.io.IOException Thrown if a problem occurs in the file system.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException Thrown
     *         if invalid xml is found whilst parsing an input SpreadsheetML
     *         file.
     */
    private void openWorkbook(File file) throws FileNotFoundException,
            IOException, InvalidFormatException {
        FileInputStream fis = null;
        try {
//            System.out.println("Opening workbook [" + file.getName() + "]");

            fis = new FileInputStream(file);

            // Open the workbook and then create the FormulaEvaluator and
            // DataFormatter instances that will be needed to, respectively,
            // force evaluation of forumlae found in cells and create a
            // formatted String encapsulating the cells contents.
            workbook = WorkbookFactory.create(fis);
            evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            formatter = new DataFormatter(true);
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
    }

    /**
     * Called to convert the contents of the currently opened workbook into
     * a Json file.
     */
    private void convertToJson() {
        Sheet sheet = null;
        Row row = null;
        int lastRowNum = 0;
        JsonData = new ArrayList<ArrayList<String>>();

        //       System.out.println("Converting files contents to Json format for workbook ="+workbook);


        // Discover how many sheets there are in the workbook....
        int numSheets = workbook.getNumberOfSheets();
// System.out.println("got this number of sheets "+numSheets);
        // and then iterate through them.
        for (int i = 0; i < numSheets; i++) {

            // Get a reference to a sheet and check to see if it contains
            // any rows.
            sheet = workbook.getSheetAt(i);
            sheetName = sheet.getSheetName();
            startSheet(i, numSheets, sheetName);
//            System.out.println("getting sheet "+sheetName);     
            if (sheet.getPhysicalNumberOfRows() > 0) {

                // Note down the index number of the bottom-most row and
                // then iterate through all of the rows on the sheet starting
                // from the very first row - number 1 - even if it is missing.
                // Recover a reference to the row and then call another method
                // which will strip the data from the cells and build lines
                // for inclusion in the resylting Json file.
                lastRowNum = sheet.getLastRowNum();
//                System.out.println("convertToJson doing rows for lastRow "+lastRowNum);
                for (int j = 1; j <= lastRowNum; j++) {
                    row = sheet.getRow(j);
                    rowToJson(row);
                }
                endSheet(i, numSheets);
            }
        }
    }

    /**
     * Called to actually save the data recovered from the Excel workbook
     * as a Json file.
     *
     * @param file An instance of the File class that encapsulates a handle
     *             referring to the Json file.
     * @throws java.io.FileNotFoundException Thrown if the file cannot be found.
     * @throws java.io.IOException Thrown to indicate and error occurred in the
     *                             underylying file system.
     */
    private void saveJsonFile(File file)
            throws FileNotFoundException, IOException {
        FileWriter fw = null;
        BufferedWriter bw = null;
        ArrayList<String> line = null;
        StringBuffer buffer = null;
        String JsonLineElement = null;
        boolean lineIsJson = false;
        System.out.println("asJson is:");
        System.out.println(asJson());
        try {

            //           System.out.println("Saving the Json file [" + file.getName() + "]");
/*
            need to add [ to start and ] to end of output
            also need to change separator to \t for tab separated values
            no quotes to separate data values or commas
             */
            // Open a writer onto the Json file.
            fw = new FileWriter(file);
            bw = new BufferedWriter(fw);
            /*
            // Step through the elements of the ArrayList that was used to hold
            // all of the data recovered from the Excel workbooks' sheets, rows
            // and cells.
            for (int i = 0; i < JsonData.size(); i++) {
            buffer = new StringBuffer();
            
            // Get an element from the ArrayList that contains the data for
            // the workbook. This element will itself be an ArrayList
            // containing Strings and each String will hold the data recovered
            // from a single cell. The for() loop is used to recover elements
            // from this 'row' ArrayList one at a time and to write the Strings
            // away to a StringBuffer thus assembling a single line for inclusion
            // in the Json file. If a row was empty or if it was short, then
            // the ArrayList that contains it's data will also be shorter than
            // some of the others. Therefore, it is necessary to check within
            // the for loop to ensure that the ArrayList contains data to be
            // processed. If it does, then an element will be recovered and
            // appended to the StringBuffer.
            line = JsonData.get(i);
            JsonLineElement = line.get(0);
            lineIsJson = (JsonLineElement != null) && (JsonLineElement.contains("{") || JsonLineElement.contains("}") 
            || JsonLineElement.contains("[") || JsonLineElement.contains("]") ) ;
            if (!lineIsJson) {
            buffer.append("[");
            }
            for (int j = 0; j < line.size(); j++) {
            if (line.size() > j) {
            JsonLineElement = line.get(j);
            if (JsonLineElement != null) {
            if (lineIsJson)
            buffer.append(JsonLineElement);
            else
            buffer.append( "\"" + JsonLineElement +"\"");
            }
            }
            if ( !lineIsJson && (j < (line.size() - 1)) ) {
            buffer.append(separator);
            }
            
            }
            if (!lineIsJson) {
            buffer.append("]");
            }
            //System.out.println(buffer.toString());
            // Once the line is built, write it away to the Json file.
            bw.write(buffer.toString().trim());
            
            // Condition the inclusion of new line characters so as to
            // avoid an additional, superfluous, new line at the end of
            // the file.
            if (i < (JsonData.size() - 1)) {
            bw.newLine();
            }
            } */
            bw.write(asJson());
        } finally {
            if (bw != null) {
                bw.flush();
                bw.close();
            }
        }
    }

    private String asJson() {
        return sheets.toString();
    }

    private String sheetAsJson() {
        ArrayList<String> line = null;
        StringBuffer buffer = new StringBuffer();
        String jsonLine = null;
        boolean lineIsJson = false;
        for (int i = 0; i < JsonData.size(); i++) {
            line = JsonData.get(i);

            //Is this entry json object notation or is it raw row data? 
            if (line.size() > 0) {
                jsonLine = line.get(0);
            } else {
                jsonLine = null;
            }
            lineIsJson = (jsonLine != null) && (jsonLine.contains("{") || jsonLine.contains("}")
                    || jsonLine.contains("[") || jsonLine.contains("]"));

            if (!lineIsJson) {
                buffer.append("[");
            }
            for (int j = 0; j < line.size(); j++) {
                if (line.size() > j) {
                    jsonLine = line.get(j);
                    if (jsonLine != null) {
                        buffer.append(jsonLine);
                    }
                }
                if (!lineIsJson && (j < (line.size() - 1))) {
                    buffer.append("\t");
                }

            }
            if (!lineIsJson) {
                if (i < JsonData.size() - 2) {
                    buffer.append("],");
                } else {
                    buffer.append("]");
                }
            }
            buffer.append("\n"); // end the object with newline
        }
        return buffer.toString();
    }

    private String cellValue(Cell cell) {
        String result = "";    
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
           result = cell.getStringCellValue();   
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
           result = Double.toString(cell.getNumericCellValue());
           if (result.contains("E")) {
              result = result.substring(0, 1) + result.substring(2, result.indexOf("E")); 
           }
           if (result.indexOf(".0") == result.length()-2) {
              result = result.substring(0, result.indexOf(".0"));  
           }
        }   
        return result;
    }

    /**
     * Called to convert a row of cells into a line of data that can later be
     * output to the Json file.
     *
     * @param row An instance of either the HSSFRow or XSSFRow classes that
     *            encapsulates information about a row of cells recovered from
     *            an Excel workbook.
     */
    private void rowToJson(Row row) {
        Cell cell = null;
        int lastCellNum = 0;
        ArrayList<String> JsonLine = new ArrayList<String>();
        // Check to ensure that a row was recovered from the sheet as it is
        // possible that one or more rows between other populated rows could be
        // missing - blank. If the row does contain cells then...
        if (row != null) {

            // Get the index for the right most cell on the row and then
            // step along the row from left to right recovering the contents
            // of each cell, converting that into a formatted String and
            // then storing the String into the JsonLine ArrayList.
            lastCellNum = row.getLastCellNum();
//  System.out.println("rowToJson called with "+row +" and lastCell of "+lastCellNum+formatter.formatCellValue(row.getCell(0), evaluator) );
            for (int i = 0; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if (cell == null) {
                    JsonLine.add("");
                } else {
                 //   System.out.println("cell = " + cellValue(cell)); // this always gives correct raw value!
                    if (cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
                   //     JsonLine.add(formatter.formatCellValue(cell)); // the formatter has some issues...
                        JsonLine.add(cellValue(cell));
                    } else {
                        JsonLine.add(formatter.formatCellValue(cell, evaluator));
                    }
                }
            }
            // Make a note of the index number of the right most cell. This value
            // will later be used to ensure that the matrix of data in the Json file
            // is square.
            if (lastCellNum > maxRowWidth) {
                maxRowWidth = lastCellNum;
            }
        }
        JsonData.add(JsonLine);
    }

    private void startSheet(int sheetNumber, int numSheets, String sheetName) {
        ArrayList<String> JsonLine = new ArrayList<String>();
//  System.out.println("startSheet called with "+sheetNumber+" of "+numSheets+" with name "+sheetName);
        JsonLine.add("{\"workbook_name\":\"" + sheetName + "\", \"workbook_data\":[");
        JsonData.add(JsonLine);
    }

    private void endSheet(int sheetNumber, int numSheets) {
        ArrayList<String> JsonLine = new ArrayList<String>();
        if (sheetNumber < numSheets - 1) {
            JsonLine.add("]},");
        } else {
            JsonLine.add("]}");
        }
        JsonData.add(JsonLine);

        sheets.append(sheetAsJson());
        JsonData.clear();
    }
}
