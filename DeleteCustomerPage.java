package com.inetbanking_pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class DeleteCustomerPage {
	WebDriver driver;
	public DeleteCustomerPage(WebDriver driver) {
		driver.get("https://demo.guru99.com/v4/manager/DeleteCustomerInput.php");
		driver.findElement(By.name("cusid")).sendKeys("mngr26593");
		   driver.findElement(By.cssSelector("button[type='submit']")).click();
		   driver.findElement(By.cssSelector("button[type='reset']")).click();	// TODO Auto-generated constructor stub
	}

}
