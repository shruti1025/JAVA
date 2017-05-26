package com.rbc.tlm.dbconnection;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class LoadConfiguration {
	public String getSQLQuery(String process){
		Properties properties = new Properties();
		String fileName = "query.sql";
	    try {
	    	InputStream inpStream = new FileInputStream(fileName);
			properties.load(inpStream);
		} catch (IOException ioException) {
			ioException.printStackTrace();
		}
	    System.out.println(properties.getProperty("Operator"));
	    return properties.getProperty(process);
	  }
}
