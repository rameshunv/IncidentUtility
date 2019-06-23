package com.mphasis.incident.incidentUtility;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.util.HashMap;
import java.util.Map;


public class TicketOwnerExcelReader {

    public  static Map<String, String> loadTicketOwnerDataMap(String filePath) throws  Exception{
        Map<String, String> ticketOwnerDataMap= new HashMap<>();
        Workbook workbook = WorkbookFactory.create(new File(filePath));
        Sheet sheet = workbook.getSheetAt(0);
        sheet.removeRow(sheet.getRow(0));
        sheet.forEach(row -> {
            if(row.getCell(0).getRichStringCellValue()!= null)
            ticketOwnerDataMap.put(row.getCell(0).getRichStringCellValue().getString().trim(),row.getCell(3).getRichStringCellValue().getString());
        });
        workbook.close();
        return ticketOwnerDataMap;
    }

/*
    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        *//*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        *//*

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach wih lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        *//*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        *//*

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

       *//* // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }*//*

        // 3. Or you can use Java 8 forEach loop with lambda
        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.removeRow(sheet.getRow(0));
        Map<String, String> masterServerList= new HashMap<>();
        sheet.forEach(row -> {

            System.out.print(row.getCell(0).getRichStringCellValue().getString() + "\t"+row.getCell(3).getRichStringCellValue().getString() );
            masterServerList.put(row.getCell(0).getRichStringCellValue().getString(),row.getCell(3).getRichStringCellValue().getString());
              //  printCellValue(cell);

            System.out.println();
        });
        System.out.println("Size>>>>>>>"+masterServerList.size());
        // Closing the workbook
        workbook.close();
    }

    private static void printCellValue(Cell cell) {
        System.out.print("<><><><><><><><><><><><><><><><"+cell.getRichStringCellValue().getString());
        *//*switch (cell.getCellTypeEnum()) {
             case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;

            case FORMULA:
                System.out.print("FPRMULAA>>>>>>>>>>>>>>>>>>"+cell.getCellFormula());
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");*//*
    }*/
}
