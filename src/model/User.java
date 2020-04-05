package model;

public class User {
	private String mobileno;
	private String empname;
	
	
	public User(String mobileno, String empname) {
		super();
		this.mobileno = mobileno;
		this.empname = empname;
	}
	public String getMobileno() {
		return mobileno;
	}
	public void setMobileno(String mobileno) {
		this.mobileno = mobileno;
	}     
	public String getEmpname() {
		return empname;
	}
	public void setCustomername(String empname) {
		this.empname = empname;
	}
	
	
}
