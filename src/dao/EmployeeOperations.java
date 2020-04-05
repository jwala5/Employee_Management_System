package dao;
import java.util.*;
import model.*;


public class EmployeeOperations {
	Excelfile excel=new Excelfile();
	static List<Employee> list1 = new ArrayList<Employee>();
	
		public List<Employee> insert(Employee employee) throws Exception {
			list1.add(employee);
			excel.excelgenerator(employee,list1); 
			list1.remove(employee);
			return list1;
		}
		public void getAllInstrument(){
			excel.excelreader("Empsheet.xlsx");
		}
	}


