package com.nickjmoss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelEditor {
    // Get the excel workbook from the file path and return it
    public static Workbook getWorkbook(String excelFilePath) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(excelFilePath));
            workbook.close();
            return workbook;
        } catch (IOException e) {
            System.out.println(e);
        }
        return workbook;
    }

    // Function to create and return the name of the new spreasheet file
    // the file is named in this format:
    // path/fileName-DECOMPRESSED-DATE.xlsx
    public static String outputFile(String fileName, String DorC) {
        String str = fileName;
        Integer index = str.lastIndexOf("/");
        String prefix = str.substring(0, index + 1);
        String suffix = str.substring(index + 1, str.length());
        Integer fileSplice = suffix.lastIndexOf(".");
        String date = LocalDate.now().toString().replace(' ', '-');
        String newSuffix = suffix.substring(0, fileSplice) + "-" + DorC + "-" + date
                + suffix.substring(fileSplice, suffix.length());

        return prefix + newSuffix;
    }

    public static Sheet setColumnHeaders(Sheet inputSheet, Sheet outputSheet, String DorC) {

        int cellCount = 0;

        Row inputHeader = inputSheet.getRow(0);

        Row outputHeader = outputSheet.createRow(0);

        for (Cell cell : inputHeader) {
            Cell newCell = outputHeader.createCell(cellCount);

            String value = cell.getStringCellValue();

            if (DorC == "DECOMPRESSED") {
                if (value.equalsIgnoreCase("first_image")) {
                    newCell.setCellValue("IMAGE_NBR");
                    cellCount++;
                } else if (value.equalsIgnoreCase("last_image")) {
                    continue;
                } else {
                    newCell.setCellValue(value);
                    cellCount++;
                }
            } else {
                if (value.equalsIgnoreCase("image")) {
                    newCell.setCellValue("FIRST_IMAGE");
                    outputHeader.createCell(cellCount + 1).setCellValue("LAST_IMAGE");
                    cellCount += 2;
                    continue;
                } else {
                    newCell.setCellValue(value);
                    cellCount++;
                }
            }
        }
        return outputSheet;
    }

    public String decompress(String file) {
        /*
         * Decompress Function:
         * 
         * Parameters: Takes the existing path of a collapsed/compressed excel
         * spreadsheet as a string. Ex: "/user/projects/spreadsheet.xlsx"
         * 
         * This function takes a collapsed excel spreadsheet and expands each row based
         * on how many images there are with the same record
         * information. For example, if there is record that is displayed as such within
         * the original spreadsheet:
         * PROJECT_ID DGS FIRST_IMAGE LAST_IMAGE NAME ETC
         * 12345 12345 1 20 Record Example
         * 
         * The decompress function will expand this record so that each record image
         * gets its own row like such:
         * PROJECT_ID DGS IMAGE_NBR NAME ETC
         * 12345 12345 1 Record Example
         * 12345 12345 2 Record Example
         * 12345 12345 3 Record Example
         * .
         * .
         * .
         * 12345 12345 20 Record Example
         * 
         * This will all be written to a new spreadsheet exported and saved in the same
         * directory as the original spreadsheet.
         */

        // Try-Catch block since we are dealing with files
        try {

            // Create Output File name
            String outputFileName = outputFile(file, "DECOMPRESSED");

            // Create the new Excel file we are going to write to
            // and export
            Workbook newExcelFile = new XSSFWorkbook();

            // Create sheet in that Excel file and return the sheet
            // to the variable new sheet
            newExcelFile.createSheet("DECOMPRESSED");

            Sheet newSheet = newExcelFile.getSheetAt(0);

            // creating Excel Workbook and Sheet instance
            // from the the file provided. This is the original
            // spreasheet
            Workbook wb = getWorkbook(file);

            if (wb == null) {
                newExcelFile.close();
                throw new IOException("You must provide an Excel file with the extension .xlsx");
            }

            // Get the first/only sheet and store its reference in 'sheet'
            Sheet sheet = wb.getSheetAt(0);

            if (sheet.getLastRowNum() < 1) {
                newExcelFile.close();
                throw new IOException(
                        "The file you provided does not have sufficient rows to work with. The file must have at least one row of info other than the column headings.");
            }

            // Create a new data formatter to format cell values
            DataFormatter formatter = new DataFormatter();

            String testValue = formatter.formatCellValue(sheet.getRow(0).getCell(1));
            if (testValue.equalsIgnoreCase("Image")) {
                newExcelFile.close();
                throw new IOException(
                        "It appears that the file you provided is already decompressed, did you mean to compress this file?");
            }
            ;
            // Set the Column Headers for the new spreadsheet
            newSheet = setColumnHeaders(sheet, newSheet, "DECOMPRESSED");

            // rowCount will be used to create the proper row number in the
            // new spreadsheet as we iterate through the original spreadsheet
            int rowCount = 1;

            // Using the Java Iterator class, create an iterator to iterate
            // through each row in the sheet
            Iterator<Row> itr = sheet.iterator();

            // Skip over the first row since that contains the column headings
            // which we already used when we set the column headings for the new
            // spreadsheet
            if (itr.hasNext()) {
                itr.next();
            }

            while (itr.hasNext()) {
                // Get the next row from input file
                Row row = itr.next();

                // Array that contains the value of each cell in the row as an
                // element of the array.
                ArrayList<String> rowArray = new ArrayList<String>();

                // Create iterator that will traverse each cell in the row
                // and add it's value to the rowArray
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    // Get next cell from input file
                    Cell cell = cellIterator.next();

                    // Test whether the type of data in the cell is a String or Number
                    switch (cell.getCellType()) {
                        case STRING:
                            // If the value is a String, get the value in the cell and store it in 'value'.
                            String value = cell.getStringCellValue();
                            rowArray.add(value);
                            break;

                        case NUMERIC: // field that represents number cell type
                            // Write to the new spreadsheet
                            String numValue = String.valueOf((int) cell.getNumericCellValue());
                            rowArray.add(numValue);
                            break;

                        default:
                    }
                }

                // loop_start = the value of FIRST_IMAGE
                int loop_start = Integer.valueOf(rowArray.get(2));
                // loop_end = the value of LAST_IMAGE
                int loop_end = Integer.valueOf(rowArray.remove(3));

                // Loop to create a row and increment loop_start by 1
                while (loop_start <= loop_end) {

                    // create new row, row number is equal to rowCount
                    Row newRow = newSheet.createRow(rowCount);

                    // variable to keep track of the cell number
                    int cellCount = 0;

                    // Loop through rowArray which is the array that was created earlier.
                    for (String info : rowArray) {

                        // If the index of the cell is 2, then set the cell's value to loop_start
                        // because this is the IMAGE_NBR column.
                        if (cellCount == 2) {
                            // Create new cell
                            Cell newCell = newRow.createCell(cellCount);
                            // Set cell value to loop_start
                            newCell.setCellValue(String.valueOf(loop_start));
                            // Increment cellCount by 1
                            cellCount++;
                        } else {
                            // Create new cell
                            Cell newCell = newRow.createCell(cellCount);
                            // Set cell value to current element in rowArray
                            newCell.setCellValue(info);
                            // Increment cellCount by 1
                            cellCount++;
                        }
                    }
                    // Increment rowCount by 1
                    rowCount++;
                    // Increment loop_start by 1
                    loop_start++;

                }

            }
            // Close the original excel file
            wb.close();

            // Create new file output with the outputFileName that was created previously
            FileOutputStream output = new FileOutputStream(new File(outputFileName));

            // Write newExcelFile to the output file
            newExcelFile.write(output);

            // Close the newExcelFile
            newExcelFile.close();

            return "The file was decompressed successfully! Your file is located at " + outputFileName;
        } catch (Exception e) {
            return e.getMessage();
        }
    }

    public static Boolean validateRows(Row row1, Row row2) {
        /*
         * This function is called within the compress function.
         * It takes two Row objects from the same Excel sheet as input
         * and compares them to see if their info is the same. This function
         * compares every column except the IMAGE_NBR column because each row
         * has a unique image number.
         * 
         * If the rows are the same it returns true, if they are not it returns false.
         */

        // Get number of cells
        int numOfCells = row1.getPhysicalNumberOfCells();
        // Cell index
        int i = 0;
        // Default the result to true
        Boolean result = true;
        // Use the Apacha POI DataFormatter to get cell values
        // Ex: String cellValue = formatter.formatCellValue(cell) -- This will convert
        // the value to a string
        // Ex: int cellValue = formatter.formatCellValue(cell) -- This will convert the
        // value to an integer
        DataFormatter formatter = new DataFormatter();

        // Loop through the cells as long as the cellIndex is less than the numOfCells
        while (i < numOfCells) {
            // If the cell index is 1 this means that it is the IMAGE_NBR column so
            // we will skip the validation on this column
            if (i == 1) {
                i++;
                continue;
            }

            // Save each cell in a variable
            Cell cell1 = row1.getCell(i);
            Cell cell2 = row2.getCell(i);

            // Save the value of each cell using DataFormatter
            String cell1Value = formatter.formatCellValue(cell1);
            String cell2Value = formatter.formatCellValue(cell2);

            // If the cell values are the same for both rows then
            // continue the loop else set result to false and break
            // out of the loop
            if ((cell1Value.equals(cell2Value))) {
                i++;
                continue;
            } else {
                result = false;
                break;
            }
        }
        return result;
    }

    public void lastRowSame(Row row1, Row row2, Sheet sheet, String first_image, String last_image, int rowCount) {
        // Create a new row
        Row newRow = sheet.createRow(rowCount);

        // Set the last image number to the IMAGE_NBR of the current row1
        last_image = String.valueOf((int) row2.getCell(1).getNumericCellValue());

        // index of the cell
        int cellCount = 0;

        DataFormatter formatter = new DataFormatter();

        // Create iterator that will traverse each cell in the row
        // and add it's value to the rowArray
        Iterator<Cell> cellIterator = row2.cellIterator();

        while (cellIterator.hasNext()) {
            // Get next cell from input file
            Cell cell = cellIterator.next();

            // Format the cell value and store it as a String in cellValue
            String cellValue = formatter.formatCellValue(cell);

            if (cellCount == 1) {
                Cell newCell = newRow.createCell(cellCount);

                newCell.setCellValue(first_image);

                cellCount++;

                newCell = newRow.createCell(cellCount);

                newCell.setCellValue(last_image);

                cellCount++;
            } else {
                Cell newCell = newRow.createCell(cellCount);

                newCell.setCellValue(cellValue);

                cellCount++;
            }
        }
    }

    public void lastRowDiff(Row row1, Row row2, Sheet sheet, String first_image, String last_image, int rowCount) {

        // Create a new row
        Row newRow = sheet.createRow(rowCount);

        // Set the last image number to the IMAGE_NBR of the current row1
        last_image = String.valueOf((int) row1.getCell(1).getNumericCellValue());

        // index of the cell
        int cellCount = 0;

        DataFormatter formatter = new DataFormatter();

        // Create iterator that will traverse each cell in the row
        // and add it's value to the rowArray
        Iterator<Cell> cellIterator1 = row1.cellIterator();
        Iterator<Cell> cellIterator2 = row2.cellIterator();

        while (cellIterator1.hasNext()) {
            // Get next cell from input file
            Cell cell = cellIterator1.next();

            // Format the cell value and store it as a String in cellValue
            String cellValue = formatter.formatCellValue(cell);

            if (cellCount == 1) {
                Cell newCell = newRow.createCell(cellCount);

                newCell.setCellValue(first_image);

                cellCount++;

                newCell = newRow.createCell(cellCount);

                newCell.setCellValue(last_image);

                cellCount++;
            } else {
                Cell newCell = newRow.createCell(cellCount);

                newCell.setCellValue(cellValue);

                cellCount++;
            }
        }
        cellCount = 0;
        rowCount++;

        // Create a new row
        newRow = sheet.createRow(rowCount);

        while (cellIterator2.hasNext()) {
            // Get next cell from input file
            Cell cell = cellIterator2.next();

            // Format the cell value and store it as a String in cellValue
            String cellValue = formatter.formatCellValue(cell);

            if (cellCount == 1) {
                Cell newCell = newRow.createCell(cellCount);

                String first_and_last = formatter.formatCellValue(row2.getCell(cellCount));

                newCell.setCellValue(first_and_last);

                cellCount++;

                newCell = newRow.createCell(cellCount);

                newCell.setCellValue(first_and_last);

                cellCount++;
            } else {
                Cell newCell = newRow.createCell(cellCount);

                newCell.setCellValue(cellValue);

                cellCount++;
            }
        }

    }

    public String compress(String file) {
        /*
         * Compress Function:
         * 
         * Parameters: Takes the existing path of a decompressed/exploded excel
         * spreadsheet as a string. Ex: "/user/projects/spreadsheet.xlsx"
         * 
         * This function takes an exploded excel spreadsheet and collapses each row
         * based on how many rows there are with the same
         * information. For example, if there are records that are displayed as such
         * within the original spreadsheet:
         * PROJECT_ID DGS IMAGE_NBR NAME ETC
         * 12345 12345 1 Record Example
         * 12345 12345 2 Record Example
         * 12345 12345 3 Record Example
         * .
         * .
         * .
         * 12345 12345 20 Record Example
         * 
         * The decompress function will compress this record so that all the similar
         * records are in one row with two columns added to it
         * FIRST_IMAGE and LAST_IMAGE:
         * PROJECT_ID DGS FIRST_IMAGE LAST_IMAGE NAME ETC
         * 12345 12345 1 20 Record Example
         * 
         * This will all be written to a new spreadsheet exported and saved in the same
         * directory as the original spreadsheet.
         * 
         * TO-DO:
         * Create a validating function that takes two rows as input and returns if the
         * rows are the same except for their
         * IMAGE_NBR.
         * 
         */
        try {

            // Create Output File name
            String outputFileName = outputFile(file, "COMPRESSED");

            // Create the new Excel file we are going to write to
            // and export
            Workbook newExcelFile = new XSSFWorkbook();

            // Create sheet in that Excel file and return the sheet
            // to the variable new sheet
            newExcelFile.createSheet("COMPRESSED");
            Sheet newSheet = newExcelFile.getSheetAt(0);

            // creating Excel Workbook and Sheet instance
            // from the the file provided. This is the original
            // spreasheet
            Workbook wb = getWorkbook(file);

            if (wb == null) {
                newExcelFile.close();
                throw new IOException("You must provide an Excel file with the extension .xlsx");
            }

            // Get the first/only sheet and store its reference in 'sheet'
            Sheet sheet = wb.getSheetAt(0);

            if (sheet.getLastRowNum() < 1) {
                newExcelFile.close();
                throw new IOException(
                        "The file you provided does not have sufficient rows to work with. The file must have at least one row of info other than the column headings.");
            }

            // Set the Column Headers for the new spreadsheet
            newSheet = setColumnHeaders(sheet, newSheet, "COMPRESSED");
            // Create a new DataFormatter to format the cell values
            DataFormatter formatter = new DataFormatter();

            String testValue = formatter.formatCellValue(sheet.getRow(0).getCell(2));
            if (testValue.equalsIgnoreCase("first_image")) {
                newExcelFile.close();
                throw new IOException(
                        "It appears that the file you provided is already compressed, did you mean to decompress this file?");
            }
            ;

            // Using the Java Iterator class, create two iterators to iterate
            // through each row in the sheet, one interator starting at the first
            // row and the second iterator starting at the second row.
            Iterator<Row> itr1 = sheet.iterator();
            Iterator<Row> itr2 = sheet.iterator();

            // Start the first iterator at the FIRST ROW, not at
            // at the column headings row
            if (itr1.hasNext()) {
                itr1.next();
            }

            // Start the second iterator at the SECOND ROW
            if (itr2.hasNext()) {
                itr2.next();
                itr2.next();
            }

            // Set the first and last image numbers to null
            String first_image = null;
            String last_image = null;

            // We need to track the loops so that we know when we are dealing
            // with a new set of identical rows
            int loop_tracker = 0;

            // rowCount will be used to create the proper row number in the
            // new spreadsheet as we iterate through the original spreadsheet
            int rowCount = 1;

            while (itr1.hasNext() && itr2.hasNext()) {
                // Get the next row from input file
                Row row1 = itr1.next();
                Row row2 = itr2.next();

                // If this is the first loop for this set of rows then
                // set the first_image number to the IMAGE_NUMBER for row1
                if (loop_tracker == 0) {
                    first_image = String.valueOf((int) row1.getCell(1).getNumericCellValue());
                }

                // An array to store all of the values that are the same for this set
                // of rows.
                ArrayList<String> compressedArray = new ArrayList<String>();

                // Use the validateRows function to test if the rows are the same.
                if (validateRows(row1, row2)) {
                    if (row2.getRowNum() == sheet.getLastRowNum()) {
                        lastRowSame(row1, row2, newSheet, first_image, last_image, rowCount);
                        break;
                    }
                    // If they are the same, loop through again and test the next set of rows
                    loop_tracker++;
                    continue;
                } else {
                    if (row2.getRowNum() == sheet.getLastRowNum()) {
                        lastRowDiff(row1, row2, newSheet, first_image, last_image, rowCount);
                        break;
                    }
                    // If the rows are different, this means row1 is different from row2 and all
                    // of the rows preceding row1 are identical to row1 so we only need to store
                    // row1's info in compressedArray.

                    // Set the last image number to the IMAGE_NBR of the current row1
                    last_image = String.valueOf((int) row1.getCell(1).getNumericCellValue());
                    // Reset the loop tracker
                    loop_tracker = 0;

                    // Create iterator that will traverse each cell in the row
                    // and add it's value to the rowArray
                    Iterator<Cell> cellIterator1 = row1.cellIterator();
                    while (cellIterator1.hasNext()) {
                        // Get next cell from input file
                        Cell cell = cellIterator1.next();

                        // Format the cell value and store it as a String in cellValue
                        String cellValue = formatter.formatCellValue(cell);

                        // Add the value to the array
                        compressedArray.add(cellValue);
                    }

                    // index of the cell
                    int cellCount = 0;

                    // Create a new row
                    Row newRow = newSheet.createRow(rowCount);

                    // Loop through the compressedArray
                    for (String cellValue : compressedArray) {
                        // If the cell's index is 1 then set the newCell's value to
                        // first_image and then set the next cell in the row to last_image
                        if (cellCount == 1) {
                            Cell newCell = newRow.createCell(cellCount);

                            newCell.setCellValue(first_image);

                            cellCount++;

                            newCell = newRow.createCell(cellCount);

                            newCell.setCellValue(last_image);

                            cellCount++;
                        } else {
                            Cell newCell = newRow.createCell(cellCount);

                            newCell.setCellValue(cellValue);

                            cellCount++;
                        }

                    }
                }
                rowCount++;
            }

            // Close the original excel file
            wb.close();

            // Create new file output with the outputFileName that was created previously
            FileOutputStream output = new FileOutputStream(new File(outputFileName));

            // Write newExcelFile to the output file
            newExcelFile.write(output);

            // Close the newExcelFile
            newExcelFile.close();

            return "The file was compressed successfully and is located at " + outputFileName;
        } catch (Exception e) {
            return e.getMessage();
        }
    }
};
