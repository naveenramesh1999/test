package Excel;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Scanner;

public class ExcelReadAndWrite {
    FileInputStream inputStream;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFRow row;
    XSSFCell cell;
    int numberOfRows ,numberOfCells;
    String home = System.getProperty("user.dir");
    String filePath = home + File.separator + "sample.xlsx";

    public void write() throws IOException {
        Scanner scanner = new Scanner(System.in);
        File file = new File(filePath);
        inputStream = new FileInputStream(filePath);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.createSheet("sheet2");

        System.out.println("Enter the number of Rows");
        numberOfRows = scanner.nextInt();
        System.out.println("Enter the number of cell");
        numberOfCells = scanner.nextInt();
        for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
            row = sheet.createRow(rowIndex);

            for (int cellIndex = 0; cellIndex < numberOfCells; cellIndex++) {
                cell = row.createCell(cellIndex);
                System.out.println("Enter the names");
                cell.setCellValue(scanner.next());

            }
        }
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        outputStream.close();
    }

    public void read() throws IOException {

        inputStream = new FileInputStream(filePath);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet("sheet2");

        for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
            row = sheet.getRow(rowIndex);

            for (int cellIndex = 0; cellIndex < numberOfCells; cellIndex++) {
                cell = row.getCell(cellIndex);

                String text = cell.getStringCellValue();
                System.out.println(text);
            }
        }

    }

    public static void main(String[] args) throws IOException {
        ExcelReadAndWrite excelReadAndWrite = new ExcelReadAndWrite();
        excelReadAndWrite.write();
        excelReadAndWrite.read();
    }
}
