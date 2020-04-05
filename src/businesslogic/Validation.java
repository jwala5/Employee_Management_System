package businesslogic;

public class Validation{
	public boolean adminlogin(String username,String password) {
		
		if(username.equals("jwala@") && password.equals("00000"))
				return true;
		else 
				return false;
	}
}