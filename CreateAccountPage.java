package com.inetbanking_pageobjects;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class CreateAccountPage {
	WebDriver driver;
	public CreateAccountPage(WebDriver driver) {
		driver.get("https://demo.guru99.com/v4/manager/addcustomerpage.php");
		
		driver.findElement(By.name("name")).sendKeys("pranali rathore");
		 driver.findElement(By.name("rad1")).click();
         driver.findElement(By.name("dob")).sendKeys("01-06-1993"); 
        driver.findElement(By.name("addr")).sendKeys("kothrud"); 
        driver.findElement(By.name("city")).sendKeys("pune");
        driver.findElement(By.name("state")).sendKeys("Maharashtra");
        driver.findElement(By.name("pinno")).sendKeys("123456");
        driver.findElement(By.name("telephoneno")).sendKeys("9876543210");
        driver.findElement(By.name("emailid")).sendKeys("abc@gmail.com");
        driver.findElement(By.name("password")).sendKeys("getout from here@123");
        
       
       
  
       
        driver.findElement(By.cssSelector("button[type='submit']")).click(); 
          
           Alert alert = driver.switchTo().alert(); 
        String alertText = alert.getText(); 
  
        // Check if the alert is open 
        if (alertText != null)  
        { 
            System.out.println("Alert is open, and the content is: " + alertText); 
            // You can accept the alert if needed: alert.accept(); 
        }  
        else 
        { 
            System.out.println("No alert found"); 
        }       
        driver.quit(); 
    } 
}
		

