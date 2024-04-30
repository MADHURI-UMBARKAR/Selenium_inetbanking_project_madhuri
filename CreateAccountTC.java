package com.inetbanking_testcases;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

import com.inetbanking_testbase.TestBase;

public class CreateAccountTC extends TestBase{
	
	@FindBy(linkText="New Customer")
	WebElement newCustomerLink;
	
}
