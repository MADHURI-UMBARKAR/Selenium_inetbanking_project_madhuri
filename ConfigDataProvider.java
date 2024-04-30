package com.inetbanking_utility;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

public class ConfigDataProvider {
	public Properties prop;
	public ConfigDataProvider(String fileName) {

	
		try {
			File fs=new File("./Config/"+fileName+".properties");
			FileInputStream fins=new FileInputStream(fs);
			prop=new Properties();
			prop.load(fins);
			}catch(Exception e) {
				e.printStackTrace();
				}
	}
	
	public String searchKey(String key) {
		return prop.getProperty(key);
		}

	public String getUserName() {
		return prop.getProperty("username");
	}
	public String getPassword() {
		return prop.getProperty("password");
	}
	public String getAppURL() {
		return prop.getProperty("appurl");
	}
}
