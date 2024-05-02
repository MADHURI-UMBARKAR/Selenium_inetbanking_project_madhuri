package com.inetbanking_testbase;

import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;



public class DDTesting extends TestBase {
	
	
	@DataProvider(name="LoginTestData")
	public String[][] loginTestData()
	{
		String[][]data=excelDataProvider.getExcelTestDate("ddtesting");
		return data;
	}
	
	@Test(dataProvider="LoginTestData")
	public void loginDDTest(String username,String password,String pageTitle)
	{
		lp.login(username,password);
		Assert.assertEquals(driver.getTitle(),pageTitle);
	}
	

}
