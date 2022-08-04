package Selenium;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Google {

	
	public static void main(String [] args) throws InterruptedException {
		
		System.setProperty("webdriver.chrome.driver","C:\\Users\\naveenkumar.r\\Desktop\\ChromeDrive\\chromedriver.exe");

		WebDriver webDriver = new ChromeDriver();
		webDriver.get("https://www.google.com/");
        webDriver.findElement(By.name("q")).sendKeys("Atmecs"+ Keys.ENTER);
        Thread.sleep(10000);
        webDriver.close();
	}
	
}
