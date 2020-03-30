  package com.cts.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class HomePage {
	private  By waitLoc = By.xpath("//a[@title='My Account']");
	private  By myAccLocator = By.xpath("//a[@title='My Account']");
	private  By loginLoc = By.linkText("Login");
	private  WebDriver driver;
	public HomePage(WebDriver driver) {
		this.driver=driver;
	}
	//explicitwait
	public  void explicitWait() {
		WebDriverWait wait = new WebDriverWait(driver,50);
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(waitLoc));
		
	}
	//click on myaccount
	public  void clickOnMyAccount() 
	{
		
       driver.findElement(myAccLocator).click();
		
	}
	//click on login
	public  void clickOnLogin() 
	{
		
	       driver.findElement(loginLoc).click();
	}
	
}
