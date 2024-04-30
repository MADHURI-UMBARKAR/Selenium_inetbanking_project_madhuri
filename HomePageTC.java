package com.inetbanking_testcases;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.inetbanking_testbase.TestBase;



public class HomePageTC extends TestBase {
	
	@Test(priority=1)
	public void VerifyUserLandedOnHomePage() {
		String expmangText="Welcome To Manager's Page of Guru99 Bank";
		Assert.assertEquals(hp.verifyManagerLogin(), expmangText);
	}
	
	@Test(priority=2)
	public void validateManagerId() {
		Assert.assertEquals(hp.getManagerId(), "Manager Id: mngr26593");
	}

}

