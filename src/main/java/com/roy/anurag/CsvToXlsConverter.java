package com.roy.anurag;

import com.opencsv.CSVReader;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.Scanner;

public class CsvToXlsConverter {

    public static void main(String[] args) throws IOException {

        Scanner sc = new Scanner(System.in);

        System.out.println("Enter the path to CSV File");
        String csvPath = sc.nextLine();

        FileReader csvFileReader = null;

        try {
            csvFileReader = new FileReader(csvPath);
        } catch (FileNotFoundException e) {
            System.out.println("Sorry, file not found.");
            System.exit(1);
        }

        CSVReader csvReader = new CSVReader(csvFileReader);

        String[] record;

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("Sheet1");

        int numberOfRows = 0;

        try {

            while ((record = csvReader.readNext()) != null) {
                Row row = sheet.createRow(numberOfRows++);

                int numberOfColumns = 0;

                for (String cellValue : record) {
                    Cell cell = row.createCell(numberOfColumns++);
                    cell.setCellValue(cellValue);
                }
            }

        } catch (IOException ioe) {
            System.out.println("Some problem in CSV format. Sorry request could not be processed.");
            System.exit(2);
        }


        System.out.println("Enter the path to save your XLS File");
        String xlsPath = sc.nextLine();

        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(xlsPath);
        } catch (FileNotFoundException e) {
            System.out.println("Couldn't find the specified path. Creating Sample.xls in your system's temporary directory.");
            fos = new FileOutputStream(System.getProperty("java.io.tmpdir") + "Sample.xls");
        }

        try (BufferedOutputStream bos = new BufferedOutputStream(fos)){
            wb.write(bos);
            bos.flush();
        } catch (IOException ioe) {
            System.out.println("Error while trying to write xls file. Please try again.");
        }

        fos.close();
        csvFileReader.close();
        csvReader.close();
        sc.close();

    }
}
