package AVIS.TestScripts;

import java.io.*;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;

import java.awt.Robot;
import java.awt.event.KeyEvent;

import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;
/* '#############################################################################################################################
* '## SCRIPT NAME: GUI_CancelReservation_AVIS '## BRAND: AVIS '## DESCRIPTION:
* Cancel Reservation for already created Reservations
* products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
* All the required Test Data should be available in Test Data Sheet. '##
* OUTPUT: Reservation should be created successfully.
* 
* Date Created: 19 Sep 2018
* HISTORY 05-SEP-2018 - GUIFunctions class created for GUI Common
* functionalities and CR functionality
* '#############################################################################################################################
**/
public class AVIS_GUI_CancelReservation
{

	public void clickRateshopSearchBtn(ChromeDriver driver)
	{
		JavascriptExecutor jse = (JavascriptExecutor) driver;
		String clickSearchJS = "document.getElementById('searchCommandLinkResRateCode').click()";
		jse.executeScript(clickSearchJS);
	}
	
	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport() {

		extent = Extentmanager.GetExtent();
		//test = extent.createTest("GUI");

	}
	
	//public static void main(String[] args) throws Exception {
	@Test
	public void test() throws Exception {
		// Read input from Excel
		try {
			
		for (int k=1; k<=4;k++)
		{
			
			AVIS_GUI_CancelReservation CancelAvis = new AVIS_GUI_CancelReservation();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Selenium\\Test_Data\\AVIS_GUITestData_CancelReservation.xlsx");
			String Execute = rwe.getCellData("Avis_GUI", k, 2);
			
			if (Execute.equals("Y"))
					{
				
				 System.setProperty("webdriver.chrome.driver","C:\\Selenium\\ChromeDriver\\chromedriver.exe");
				 WebDriver driver= new ChromeDriver();
				 GUIFunctions functions = new GUIFunctions(driver);
				 int a = 12;
				 driver.manage().window().maximize();
				 driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				 System.out.println(" Iteration " + k);
				 String testcasename = rwe.getCellData("Avis_GUI", k, 4);
				 String tokenURL    = rwe.getCellData("Avis_GUI", k, 6);
				 String clientURL   = rwe.getCellData("Avis_GUI", k, 7);
				 String outSTA      = rwe.getCellData("Avis_GUI", k, 8);
				 String thinClient  = clientURL+outSTA;
				 String uName       = rwe.getCellData("Avis_GUI", k, 9);
				 String pswd        = rwe.getCellData("Avis_GUI", k, 10);
				 String ResNo       = rwe.getCellData("Avis_GUI", k, 11);
				 
				 functions.openURL(thinClient);
				 functions.login(uName, pswd);
				 driver.findElement(By.xpath("//a[@data-target='#menulist\\:rateshoplink']")).click();
				 System.out.println("Clicked on Rate Shop Tab");
				 driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:searchResForm:searchString']")).sendKeys("09829756US6");
					System.out.println("Entered the existing Reservation number");
					
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:searchResForm:searchCommandLink']")).click();
					System.out.println("Clicked on the Search button");
					Thread.sleep(7000);
					driver.findElement(By.xpath("//button[@class='dropdown-toggle']//span[@class='caret']")).click();
					System.out.println("Clicked on Extras drop down Button");
					
					try {
						if (driver.findElement(By.xpath("//a[@class='rateShopCancelLinkClass ng-binding']")).isDisplayed())
							
							{
							   System.out.println("Cancel Button exists ");
							   driver.findElement(By.xpath("//a[@class='rateShopCancelLinkClass ng-binding']")).click();
							   System.out.println("Clicked on Cancel Drop down link");
							   System.out.println(driver.switchTo().alert().getText());
							   driver.switchTo().alert().accept();
							   System.out.println("Clicked on OK");
							}			
					}

					catch(Exception e)
					{
						System.out.println("Cancel Button does not exists ");
						System.out.println("The Reservation number is already Cancelled");
					}
					Thread.sleep(3000);
					System.out.println(driver.findElement(By.xpath("//form[@id='templateInfoForm']//div[@class='modal-body']")).getText());
					
					File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(src,new File("C://Selenium//Screenshots//CancelScreenshot.png"));
					
					/*
					 * to print QV data in Excel
					 */

					WebElement res = driver.findElement(By.cssSelector(
							"#quickViewPanel > div.panel-body > table > tbody > tr:nth-child(8) > td > div:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > a > span"));

					try {
						if (res.getText().isEmpty()) {
							WebDriverWait wait1 = new WebDriverWait(driver, 20);
							wait1.until(ExpectedConditions.visibilityOf(res));
						}
					} catch (Exception e2) {
						e2.printStackTrace();
					}
					WebElement table = driver.findElement(By.id("resQvForm"));
					ArrayList<WebElement> rows = (ArrayList<WebElement>) table.findElements(By.tagName("tr"));
					for (int i = 1; i < rows.size(); i++) {
						ArrayList<WebElement> cells = (ArrayList<WebElement>) rows.get(i).findElements(By.tagName("td"));
						for (int j = 0; j < cells.size(); j++) {
							String val = cells.get(j).getText();
							if (val.isEmpty()) {
								break;
							} else {
								if (j == 2) {
									val = val.replaceAll("[*]", ""); // Remover '*' before rates
									rwe.setCellData("Avis_GUI", k, a, val);
									a++;
								}
							}
						}
					}
					/*
					 * Log out and close tabs
					 */

					functions.logout();
					Thread.sleep(1000);
					functions.closeWindows();
					test = extent.createTest(testcasename);
					
					if (rwe.getCellData("Avis_GUI", k, 12).isEmpty())
					{
						rwe.setCellData("Avis_GUI", k, 14, "FAIL");
						test.log(Status.FAIL, "Fail");
					}
					else
					{
						rwe.setCellData("Avis_GUI", k, 14, "PASS");
						test.log(Status.PASS, "Pass");
					}
				}
				else
				{
					System.out.println("Execution status is N for iteration " + k + "...");
				}
				 
					}
		
		}
		finally
		{
			// TODO: handle finally clause
			extent.flush();
		}
			
		}
	}


