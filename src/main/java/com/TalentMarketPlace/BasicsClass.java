package com.TalentMarketPlace;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class BasicsClass {
	 WebDriver driver = new ChromeDriver(); /////sangita
	
	public void getData()
	{
		 //driver = new ChromeDriver() ; //namkororn
		driver.get("google");
	}
	
	public void getName()
	{
		driver.get("yahoo");
		//null pointer exception
	}
		
	}
