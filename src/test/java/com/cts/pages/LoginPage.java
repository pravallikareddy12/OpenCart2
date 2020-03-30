package com.cts.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class LoginPage {

	private  By emailLocator=By.id("input-email");
	private  By passwordLocator=By.id("input-password");
	private  By LoginLocator=By.xpath("//input[@value='Login']");
	 private  By waitLocator = By.xpath("//h2[text()='My Orders']");
	 //calling driver
	 private  WebDriver driver;
		public LoginPage(WebDriver driver) {
			this.driver=driver;
		}
//enter email
	public  void enterEmail(String email)
	{
	driver.findElement(emailLocator).sendKeys(email);
	}
//enter password
	public  void enterPassword(String password)
	{
	driver.findElement(passwordLocator).sendKeys(password);
	}
// click on login
	public  void clickOnLogin( )
	{
	driver.findElement(LoginLocator).click();
	}
// explicit wait
	public  void explicitWait() {
	WebDriverWait wait = new WebDriverWait(driver,50);
	wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(waitLocator));
	}
	//get current title
	public  String getCurrentTitle() {
		String expectedTitle = "My Account";
		String actualTitle = driver.getTitle();
		return actualTitle;
	}
	}



