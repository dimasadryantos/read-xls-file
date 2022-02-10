package main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReaderXlsFile {


    public static void main(String[] args) {
        processReadFile();
    }


    private static void processReadFile() {
        try {

            FileInputStream fileInputStream = new FileInputStream
                    (new File("/Users/dimasdz/Documents/read-xls-file/resource/PTPL-MasterDataPurchaseOrder.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object

            for (Row row : sheet) {
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellTypeEnum()) {
                        case STRING:    //field that represents string cell type
                            System.out.print(cell.getStringCellValue() + "\t\t\t");
                            break;
                        case NUMERIC:    //field that represents number cell type
                            System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            break;
                        default:
                    }
                }
                System.out.println("");
            }
        } catch (Exception exception) {
            exception.printStackTrace();
        }

    }


}

