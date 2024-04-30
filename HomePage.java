package com.inetbanking_pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class HomePage {
	WebDriver driver;
	
	@FindBy(xpath="//marquee[@class='heading3']")
	WebElement managralTxt;
	
	@FindBy(linkText="New Customer")
	WebElement newCustomerLink;
	
	
	
	
	
	public HomePage(WebDriver driver) {
		this.driver=driver;
		PageFactory.initElements(driver, this);
	}
	
	
	
	
	
	public String getManagerId()
	{
		return driver.findElement(By.xpath("//tr[@align='center']/td")).getText();
	}
	
	public String verifyManagerLogin() {
		return managralTxt.getText();
	}
	
	public  CreateAccountPage NavigateToCreateAccoutPage()
	{
		newCustomerLink.click();
		return new CreateAccountPage(driver);
	}
	

}
