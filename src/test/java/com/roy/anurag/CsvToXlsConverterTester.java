package com.roy.anurag;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.Iterator;

import static org.junit.Assert.assertEquals;


public class CsvToXlsConverterTester {

    @Test
    public void TestConversion() throws IOException, URISyntaxException {


        URL testResource1 = getClass().getResource( "/ManualXLS.xls" );
        FileInputStream manualXLS = new FileInputStream(new File(testResource1.toURI()));
        Workbook manualWB = new HSSFWorkbook(manualXLS);
        Sheet manualSheet = manualWB.getSheetAt(0);
        Iterator<Row> manualIt = manualSheet.iterator();

        URL testResource2 = getClass().getResource( "/OutputXLS.xls" );
        FileInputStream outputXLS = new FileInputStream(new File(testResource2.toURI()));
        Workbook outputWB = new HSSFWorkbook(outputXLS);
        Sheet outputSheet = outputWB.getSheetAt(0);
        Iterator<Row> outputIt = outputSheet.iterator();

        while (manualIt.hasNext()) {
            //Assert false immediately when outputXLS has less number of rows than manualXLS
            if (!outputIt.hasNext()) {
                assert(false);
            } else {
                Row oRow = outputIt.next();
                Iterator<Cell> oCellIt = oRow.iterator();
                Row mRow = manualIt.next();
                Iterator<Cell> mCellIt = mRow.iterator();

                while(mCellIt.hasNext()) {
                    Cell mCell = mCellIt.next();

                    if (oCellIt.hasNext()) {
                        Cell oCell = oCellIt.next();

                        //Compare each cell value of ManualXLS and OutputXLS
                        String mTemp = mCell.getStringCellValue();
                        String oTemp = oCell.getStringCellValue();

                        assertEquals(mTemp, oTemp);
                    } else {
                        assert(false);
                    }
                }
            }


        }

        //If outputXLS has more number of rows than manualXLS assert false
        if (outputIt.hasNext()) {
            assert(false);
        }
    }
}
