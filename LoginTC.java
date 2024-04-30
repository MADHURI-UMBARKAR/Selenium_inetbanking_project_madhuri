package com.inetbanking_testcases;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.inetbanking_pageobjects.LoginPage;
import com.inetbanking_testbase.TestBase;

public class LoginTC extends TestBase {
	
	// Will be providing implementation later

	@Test(priority = 1)
	public void verifyLoginPageTitle() {
		
		String expTitle = "Guru99 Bank Home Page";
		String actTitle = driver.getTitle();
		Assert.assertEquals(actTitle, expTitle);
	}
	
	@Test(priority = 2)
	public void verifyloginfunctionality() {
	LoginPage lp = new LoginPage(driver);
    lp.login("mngr565005", "YzAdYvU");
    
    String expTitle="Guru99 Bank Manager HomePage";
    String actTitle = driver.getTitle();
    Assert.assertEquals(actTitle, expTitle);
	}
	
}
