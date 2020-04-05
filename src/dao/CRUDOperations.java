package dao;

import java.util.*;
import model.User;

public class CRUDOperations {
	static int value = 0;
	static Map<Integer,User> map = new HashMap<Integer,User>();
	
	public Map<Integer,User> addMapUser(User user){
		
		++value;
		
		map.put(value,user);
		return map;
	}

	public Map<Integer,User> getAllMapUsers(){
		return map;
	}
}
