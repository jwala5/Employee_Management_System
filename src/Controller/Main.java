/**
 * 
 */
package Controller;
import java.util.*;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.*;


import businesslogic.Validation;
import dao.*;
import model.*;
public class Main {
	
	
//	static Map<Integer,User> map= new HashMap<Integer,User>();
//	CRUDOperations crud=new CRUDOperations();
	
	public static void main(String[] args) throws Exception
	{
		LocalDateTime myDateObj = LocalDateTime.now();
		DateTimeFormatter myFormatObj = DateTimeFormatter.ofPattern("dd-MM-yyyy HH:mm:ss");
	System.out.println("Good morning today's Date is:"+myDateObj.format(myFormatObj));	
	System.out.println("==============================================================");	
	System.out.println("                Employee Management System");
	System.out.println("==============================================================");
	System.out.println("Select your login");
	Scanner sc = new Scanner(System.in);
	FirstClass fc = new FirstClass();
//	CRUDOperations crud=new CRUDOperations();

	while(true) {
		System.out.println("1. Admin");
		System.out.println("2. Employee");
		System.out.println("3. Exit");
		System.out.println("==============================");
		String username;
		String mobileno;
		String password;
		String empname;
//		User user = null;
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		int choice = sc.nextInt();
		Validation validate = new Validation();
		switch(choice){
			case 1:
				System.out.println("Enter username: ");
				username = sc.next();
				System.out.println("Enter password: ");
				password = sc.next();
				if(validate.adminlogin(username,password) ) {
					System.out.println("LOGIN SUCCESFULL!");
					fc.admin();
				}
				else 
					System.out.println("Check your username and password");
					System.out.println("================");
				break;
			case 2:
				System.out.println("Enter your moblie number:");
//				mobileno = br.readLine();
//				System.out.print("Enter your name");
//				empname =  br.readLine();
//				user = new User(mobileno,empname);
//				crud.addMapUser(user);
//				fc.customer();
				break;
			default:
				System.out.println("Wrong Choice!!");
				break;	
			
			}
		}
	}
}