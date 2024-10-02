import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) {
        // Specify the Excel file path
        String excelFilePath = "src/assignment_1.xlsx";

        // Create an object of ExcelReaderHelper
        ExcelReaderHelper excelReader = new ExcelReaderHelper(excelFilePath);

        // Call the method to read and print the Excel file
        excelReader.readExcel();
    }
}

class Employee {
    public int emp_Id;
    String emp_Fname;
    String emp_Lname;
    String emp_Role;
    Double emp_Salary;

    Employee(int id, String Fname, String Lname, String role, Double salary) {
        this.emp_Id = id;
        this.emp_Fname = Fname;
        this.emp_Lname = Lname;
        this.emp_Role = role;
        this.emp_Salary = salary;
    }

    public void displayEmployeeDetails() {
        System.out.println("ID: " + emp_Id);
        System.out.println("First Name: " + emp_Fname);
        System.out.println("Last Name: " + emp_Lname);
        System.out.println("Role: " + emp_Role);
        System.out.println("Salary: " + emp_Salary);
    }
}

class ExcelReaderHelper {
    private String filePath;

    // Constructor
    public ExcelReaderHelper(String filePath) {
        this.filePath = filePath;
    }

    // Method to read and print the Excel content
    public void readExcel() {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Reading the first sheet
            for (Row row : sheet) {
                if (row.getRowNum() == 0)
                    continue; // Skip header row if present

                int id = (int) row.getCell(0).getNumericCellValue(); // Employee ID
                String fname = row.getCell(1).getStringCellValue(); // First Name
                String lname = row.getCell(2).getStringCellValue(); // Last Name
                String role = row.getCell(3).getStringCellValue(); // Role
                double salary = row.getCell(4).getNumericCellValue(); // Salary

                // Create an Employee object
                Employee employee = new Employee(id, fname, lname, role, salary);

                // Display employee details
                employee.displayEmployeeDetails();
                System.out.println("-------------------------");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
