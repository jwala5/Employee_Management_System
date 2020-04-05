package Controller;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import com.itextpdf.text.DocumentException;

import model.*;
import dao.*;

public class FirstClass {
	static Map<Integer,User> map= new HashMap<Integer,User>();
	//CRUDOperations crud=new CRUDOperations();
	Scanner sc = new Scanner (System.in);
	Excelfile excel= new Excelfile();
	ConverToPDF pdf =new ConverToPDF();
//Admin function
	void admin() {
		
		System.out.println("=======================");
		int choice;
		do {
		System.out.println("Choose from given options");
		System.out.println("1. Add employee");
		System.out.println("2. Update employee");
		System.out.println("3. Delete employee");
		System.out.println("4. Employee Details");
		System.out.println("5. All Employee PDF");
		System.out.println("6. Employee Salary");
		System.out.println("7. Exit");
		
		String id;
		String empname;
		String dept;
		String salary;
		String age;
		String post;
		Employee employee= null;
		
		List<Employee> list1 = new ArrayList<Employee>();
		EmployeeOperations ins = new EmployeeOperations();
		choice = sc.nextInt();
		switch(choice) {
		
		case 1:
			// enter all employee details through console
			System.out.println("Enter Employee id");
			id = sc.next();
			System.out.println("Enter Employee name");
			empname = sc.next();
			System.out.println("Enter Department");
			dept = sc.next();
			System.out.println("Enter salary");
			salary = sc.next();
			System.out.println("Enter age");
			age = sc.next();
			System.out.println("Enter post");
			post = sc.next();
			employee= new Employee(id,empname,dept,salary,age,post);
			try {
				ins.insert(employee);
			} catch (Exception e) {
			
				e.getMessage();
			}
			System.out.println("Added Successfully!");
			break;
		case 2:
			break;
		case 3: 
			// deleting employee records
			System.out.println("*----------------------------------------*");
			System.out.println("         Delete Employee Details        ");
			System.out.println("*----------------------------------------*");
			System.out.print("Enter Instrument ID You Wants to Delete : ");
			id = sc.next();
			excel.deleteitem(id);
			break;
		case 4:
			// getting employee details on console
			System.out.println("*----------------------------------------*");
			System.out.println("           Employee Details               ");
			System.out.println("*----------------------------------------*");
			System.out.println("Id         Employee_Name     Department       salary          age            post");
			System.out.println("*----------------------------------------*");
			ins.getAllInstrument();
			System.out.println("*----------------------------------------*");
			break;
		case 5:
			// converting data from excelsheet to pdf
			try {
				pdf.ConverToPDF("Empsheet.xlsx");
				
			} catch (DocumentException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
		case 6:
			break;
	default:
			System.out.println("Wrong input");
			break;
			
		}
	}while(choice!=6);
	}
	// void employee()
}