package Payless.TestScripts;

import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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

import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;
/* '#############################################################################################################################
* '## SCRIPT NAME: GUI_CancelReservation_Payles '## BRAND: PAYLESS '## DESCRIPTION:
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
public class Payless_GUI_CancelReservation
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
	
//	public static void main(String[] args) throws Exception {
	@Test
	public void test() throws Exception {
		try {
			
		// Read input from Excel
		for (int k=1; k<=10;k++)
		{
			
			//GUI_CancelReservation_Avis Avis = new GUI_CancelReservation_Avis();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Selenium\\Test_Data\\AVIS\\Payless_GUITestData_CancelReservation.xlsx");
			String Execute = rwe.getCellData("Payless_GUI", k, 2);
			
			if (Execute.equals("Y"))
			{
				
				 System.setProperty("webdriver.chrome.driver","C:\\Selenium\\ChromeDriver\\chromedriver.exe");
				 WebDriver driver= new ChromeDriver();
				 GUIFunctions functions = new GUIFunctions(driver);
				 int a = 12;
				 driver.manage().window().maximize();
				 driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				 System.out.println(" Iteration " + k);
				 String ptestcasename = rwe.getCellData("Payless_GUI", k, 4);
				// String tokenURL    = rwe.getCellData("Payless_GUI", k, 6);
				 String pclientURL   = rwe.getCellData("Payless_GUI", k, 7);
				 String poutSTA      = rwe.getCellData("Payless_GUI", k, 8);
				 String pthinClient  = pclientURL+poutSTA;
				 String puName       = rwe.getCellData("Payless_GUI", k, 9);
				 String ppswd        = rwe.getCellData("Payless_GUI", k, 10);
				 String pResNo       = rwe.getCellData("Payless_GUI", k, 11);
				 
				// functions.link(tokenURL,thinClient);
				 driver.get(pthinClient);
				 functions.login(puName, ppswd);
				 driver.navigate().refresh();
				
				 //Avis.clickRateshopSearchBtn(driver);
				 driver.findElement(By.xpath("//a[@data-target='#menulist\\:rateshoplink']")).click();
				 System.out.println("Clicked on Rate Shop Tab");
				 driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:searchResForm:searchString']")).sendKeys(pResNo);
					System.out.println("Entered the existing Reservation number");
					
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:searchResForm:searchCommandLink']")).click();
					System.out.println("Clicked on the Search button");
					Thread.sleep(6000);
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
					
					
					Thread.sleep(5000);
					//System.out.println(driver.findElement(By.xpath("//form[@id='templateInfoForm']//div[@class='modal-body']")).getText());
					//driver.findElement(By.xpath("//button[@value='OK']")).click();
					//System.out.println("Clicked on OK");
					
	/* Capture Screenshot*/
					String ScreenShotPath = "C:\\Selenium\\Screenshots\\Payless\\Payless_Cancel_Reservation\\";
					functions.ScreenCapture(ScreenShotPath , ptestcasename);
					System.out.println("The Screenshot is taken");
					
					try
					{
					   if (driver.findElement(By.xpath("//form[@id='templateInfoForm']")).isDisplayed())
					   {
						   String message = driver.findElement(By.xpath("//form[@id='templateInfoForm']//div[@class='modal-body']")).getText();
						   System.out.println(message);
						   driver.findElement(By.xpath("//div[@class='modal-footer']//div[@class='row']//div[@class='col-md-12 text-left']//button[@value='OK']")).click();
						   System.out.println("Clicked on OK");   
					   }
					}
					catch (Exception e)
					{
						System.out.println("The Reservation number is already Cancelled");
					}
					
						
					Thread.sleep(3000);
					
					//Capture QUickView Details
					WebElement res = driver.findElement(By.cssSelector("#quickViewPanel > div.panel-body > table > tbody > tr:nth-child(8) > td > div:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > a > span"));

                    try {
                           if (res.getText().isEmpty())
                           {
                                  WebDriverWait wait1 = new WebDriverWait(driver, 20);
                                  wait1.until(ExpectedConditions.visibilityOf(res));
                           }
                    } catch (Exception e2)
                    {
                           e2.printStackTrace();
                    }
                    WebElement table = driver.findElement(By.id("resQvForm"));
                    ArrayList<WebElement> rows = (ArrayList<WebElement>) table.findElements(By.tagName("tr"));
                    for (int i = 1; i < rows.size(); i++) 
                    {
                           ArrayList<WebElement> cells = (ArrayList<WebElement>) rows.get(i).findElements(By.tagName("td"));
                           for (int j = 0; j < cells.size(); j++)
                           {
                                  String val = cells.get(j).getText();
                                  if (val.isEmpty()) 
                                  {
                                         break;
                                  } else 
                                  {
                                         if (j == 2)
                                         {
                                        	 if ( a< 14)
                                        	 {
                                                val = val.replaceAll("[*]", ""); // Remover '*' before rates
                                                rwe.setCellData("Payless_GUI", k, a, val);
                                                System.out.println(val);
                                                a++;
                                        	 }
                                        	 
                                         }
                                        
                                  }
                           }
                          
                    }
                    test = extent.createTest(ptestcasename);
                    
                    if(rwe.getCellData("Payless_GUI", k, 12).isEmpty())  {
                    	
						rwe.setCellData("Payless_GUI", k, 14, "FAIL");
						test.log(Status.FAIL, "Fail");
					} else {
						rwe.setCellData("Payless_GUI", k, 14, "PASS");
						test.log(Status.PASS, "Pass");
					}

					
					/*
					 * Log out and close tabs
					 */

						
				 	functions.logout();
					Thread.sleep(2000);
					functions.closeWindows();
										
				}
				
				 
			}
		} finally {
			// TODO: handle finally clause
			extent.flush();
		}
			
			}		

	}
	
	


