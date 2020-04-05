package model;

public class Employee {

	private String id;
	private String empname;
	private String dept;
	private String salary;
	private String age;
	private String post;
	
	
	public Employee(String id, String empname, String dept, String salary, String age, String post) {
		
		this.id = id;
		this.empname = empname;
		this.dept = dept;
		this.salary = salary;
		this.age = age;
		this.post = post;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}     
	public String getSalary() {
		return salary;
	}
	public void setSalary(String salary) {
		this.salary = salary;
	}     
	public String getPost() {
		return post;
	}
	public void setPost(String post) {
		this.post = post;
	}     
	public String getAge() {
		return age;
	}
	public void setAge(String age) {
		this.age = age;
	}     
	public String getEmpname() {
		return empname;
	}
	public void setEmpname(String empname) {
		this.empname = empname;
	}
	public String getDept() {
		return dept;
	}
	public void setDept(String dept) {
		this.dept = dept;
	}
}
	