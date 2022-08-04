package Excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelWriter {

    String fileOne = "D:\\Excel\\StudentOne.txt";
    String fileTwo = "D:\\Excel\\StudentTwo.txt";
    String fileThree = "D:\\Excel\\StudentThree.txt";
    File file;
    Scanner scanner = new Scanner(System.in);
    XSSFWorkbook xssfWorkbook;


    public ArrayList<String> fileReader(File file) throws IOException {
        BufferedReader bufferedReader = new BufferedReader(new FileReader(file));

        String list = "";
        ArrayList<String> arrayList = new ArrayList<String>();
        while ((list = bufferedReader.readLine()) != null) {
            arrayList.add(list);
        }
        return arrayList;
    }
        public void writer() throws IOException {
            ArrayList<String> heading = new ArrayList<String>();
            heading.add("RollNumber");
            heading.add("Name");
            heading.add("Total");
            heading.add("Result");
            file = new File(fileOne);
            ArrayList<String> listOne = fileReader(file);
            file = new File(fileTwo);
            ArrayList<String> listTwo = fileReader(file);
            file = new File(fileThree);
            ArrayList<String> listThree = fileReader(file);

            System.out.println("Enter the file name :");
            String fileName = scanner.nextLine();
            String userDirectory = System.getProperty("user.dir");
            String path = userDirectory + File.separator + fileName;

            file = new File(path);

            FileInputStream fileInputStream = new FileInputStream(file);

            xssfWorkbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet2");

            XSSFRow xssfRow;
            XSSFCell xssfCell;

            Map<String, ArrayList> map = new TreeMap<String, ArrayList>();
            map.put("0", heading);
            map.put("1", listOne);
            map.put("2", listTwo);
            map.put("3", listThree);

            Set<String> set = map.keySet();
            int row = 0;
            
            for (String str : set) {
                xssfRow = xssfSheet.createRow(row++);


                ArrayList objArray = map.get(str);

                int cell = 0;
                for (Object obj : objArray) {
                    xssfCell = xssfRow.createCell(cell++);

                    xssfCell.setCellValue((String) obj);
                }
            }
            

            FileOutputStream fileOutputStream = new FileOutputStream(file);
            xssfWorkbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("File Successfully Created");
        }

        public static void main (String[]args) throws IOException {
            ExcelWriter excelWriter = new ExcelWriter();
            excelWriter.writer();
        }

    }
