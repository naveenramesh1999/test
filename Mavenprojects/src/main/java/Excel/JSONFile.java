package Excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;
import java.util.*;

public class JSONFile {

    String directory = System.getProperty("user.dir") + File.separator + "src" + File.separator + "Naveen.json";
    Scanner scanner = new Scanner(System.in);
    File file;

    public void writeJSON() throws IOException {

        JSONObject jsonObject = new JSONObject();
        jsonObject.put("EmployeeID", "2038");
        jsonObject.put("EmployeeName", "Naveen");
        jsonObject.put("EmployeeRole", "Trainee");
        FileWriter fileWriter = new FileWriter(directory);
        fileWriter.write(jsonObject.toJSONString());
        fileWriter.close();
        System.out.println(jsonObject);

    }

    public void readJSON() throws IOException, ParseException {

        JSONParser parser = new JSONParser();
        BufferedReader reader = new BufferedReader(new FileReader(directory));

        JSONObject object = (JSONObject) parser.parse(reader);

        String employeeId = (String) object.get("EmployeeID");
        System.out.println("Employee ID :"+employeeId);
        String employeeName = (String) object.get("EmployeeName");
        System.out.println("Employee Name :"+employeeName);
        String employeeRole = (String) object.get("EmployeeRole");
        System.out.println("Employee Role :"+employeeRole);

        System.out.println("Enter the file name :");
        String fileName = scanner.nextLine();
        String userDirectory = System.getProperty("user.dir");
        String path = userDirectory + File.separator + fileName;
        file = new File(path);

        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);

        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet3");

        XSSFRow xssfRow;
        XSSFCell xssfCell;

        ArrayList<String> header = new ArrayList<String>();
        header.add("Employee ID");
        header.add("Employee Name");
        header.add("Employee Role");

        ArrayList<String> arrayList = new ArrayList<String>();
        arrayList.add(employeeId);
        arrayList.add(employeeName);
        arrayList.add(employeeRole);

        Map<String, ArrayList> map = new TreeMap<String, ArrayList>();
        map.put("0", header);
        map.put("1", arrayList);

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

    public static void main(String[] args) throws IOException, ParseException {
        JSONFile jsonFile = new JSONFile();
        jsonFile.writeJSON();
        jsonFile.readJSON();
    }
}