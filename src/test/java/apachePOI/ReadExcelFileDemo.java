package apachePOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class ReadExcelFileDemo {

    private InputStream iStream;
    private Workbook workbook;
    private Sheet sheet;
    private Row currentRow;
    private Cell currentCell;

    public void readAndPrintExcelData(String fileName) {
        try {
            iStream = new FileInputStream(fileName);
            workbook = new XSSFWorkbook(iStream);
            sheet = workbook.getSheet("Employee_details");
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                currentRow = rowIterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    currentCell = cellIterator.next();

                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + " \t ");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + " \t ");
                    }
                }

                System.out.println();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();

        } catch (IOException ioe) {
            ioe.printStackTrace();
        } catch (Exception e) {
            System.out.println("Some Exception Occured");
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        ReadExcelFileDemo demo = new ReadExcelFileDemo();

        demo.readAndPrintExcelData(Util.READ_EXCEL_FILE);


    }

}
