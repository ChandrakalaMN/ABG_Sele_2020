package Payless.TestScripts;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.gui.report.Extentmanager;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;
/* /**
 * '#############################################################################################################################
 * '## SCRIPT NAME: GUI_CreateRentals_AVIS '## BRAND: AVIS '## DESCRIPTION:
 * Create Rentals for available Reservation
 * products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
 * All the required Test Data should be available in Test Data Sheet. '##
 * OUTPUT: Reservation should be created successfully.
 * 
 * 
 * HISTORY 05-SEP-2018 - GUIFunctions class created for GUI Common
 * functionalities and CR functionality
 * '#############################################################################################################################
 * */ 

public class Payless_GUI_CreateRentals
{	
	private static final String NULL = null;
	ExtentReports extent;
	ExtentTest test;
	
	@BeforeTest
	public void startReport()
	{

		extent = Extentmanager.GetExtent();
	
	}
	@Test
	public void test() throws Exception 
	{
		try{
			Properties prop = new Properties();
			FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\Downloads\\ABG-master\\ABG-master\\src\\AVIS\\TestData\\TestDataABGGUI.properties");
			prop.load(fis);
			//WebDriver driver;
			ChromeDriver driver = new ChromeDriver();
			GUIFunctions functions = new GUIFunctions(driver);
			System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
			driver.navigate().to(prop.getProperty("PaylessURL"));
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			Thread.sleep(2000);
			functions.txt_userid.sendKeys(prop.getProperty("USERID"));
			Thread.sleep(500);
			functions.txt_password.sendKeys(prop.getProperty("PASSWORD"));
			Thread.sleep(500);
			functions.btn_login.click();
			/*String downloadFilepath = prop.getProperty("PDFDownload");
			HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
			chromePrefs.put("profile.default_content_settings.popups", 0);
			chromePrefs.put("download.default_directory", downloadFilepath);
			ChromeOptions options = new ChromeOptions();
			options.setExperimentalOption("prefs", chromePrefs);
			DesiredCapabilities cap = DesiredCapabilities.chrome();
			cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			cap.setCapability(ChromeOptions.CAPABILITY, options);
			
			
			cap.setCapability("download.default_directory","C:");
			cap.setCapability("download.prompt_for_download","false");
			cap.setCapability("directory_upgrade","true");
			

			WebDriver driver1 = new ChromeDriver(cap);*/
				/* Login */
			Thread.sleep(3000);
		// Read input from excel
		///int intRowCount     = 100;
		for (int k = 1; k <= 100; k++)
		{
			Payless_GUI_CreateRentals avis = new Payless_GUI_CreateRentals();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Payless\\Payless_GUITestData_Rental_CheckOut.xlsx");
												
			String Execute = rwe.getCellData("Payless_GUI_Rentals", k, 2);
			
			//********Delete the files in the folder********//
			File file = new File(prop.getProperty("ScreenshotPayless"));  

			String[] myFiles;    
			if (file.isDirectory()) {
			    myFiles = file.list();
			    for (int i = 0; i < myFiles.length; i++) {
			        File myFile = new File(file, myFiles[i]); 
			        myFile.delete();
			    }
			}
			    System.out.println(" iteration " + k);
				 String TCName 	= rwe.getCellData("Payless_GUI_Rentals", k, 4);
				 String clientURL   	= rwe.getCellData("Payless_GUI_Rentals", k, 7);
				 String outSTA      	= rwe.getCellData("Payless_GUI_Rentals", k, 10);
				 String thinClient  	= clientURL+outSTA;
				 String uName       	= rwe.getCellData("Payless_GUI_Rentals", k, 8);
				 String pswd        	= rwe.getCellData("Payless_GUI_Rentals", k, 9);
				 String ResNo      		= rwe.getCellData("Payless_GUI_Rentals", k, 11);
				 String CustType		= rwe.getCellData("Payless_GUI_Rentals", k, 12);
				 String WizardNo		= rwe.getCellData("Payless_GUI_Rentals", k, 13);
				 String LastName		= rwe.getCellData("Payless_GUI_Rentals", k, 14);
				 String FirstName		= rwe.getCellData("Payless_GUI_Rentals", k, 15);
				 String COCountry 		= rwe.getCellData("Payless_GUI_Rentals", k, 16);
				 String COState			= rwe.getCellData("Payless_GUI_Rentals", k, 17);
				 String CODRLICNO 		= rwe.getCellData("Payless_GUI_Rentals", k, 18);
				 String CODOB			= rwe.getCellData("Payless_GUI_Rentals", k, 19);
				 String COCOMPANY		= rwe.getCellData("Payless_GUI_Rentals", k, 20);
				 String COADDR1			= rwe.getCellData("Payless_GUI_Rentals", k, 21);
				 String COADDR2			= rwe.getCellData("Payless_GUI_Rentals", k, 22);
				 String COADDR3			= rwe.getCellData("Payless_GUI_Rentals", k, 23);
				 String COCONTACTINFO   = rwe.getCellData("Payless_GUI_Rentals", k, 24);
				 String COEMAIL			= rwe.getCellData("Payless_GUI_Rentals", k, 25);
				 String COMOP			= rwe.getCellData("Payless_GUI_Rentals", k, 26);
				 String COCCDC			= rwe.getCellData("Payless_GUI_Rentals", k, 27);
				 String CARDNO			= rwe.getCellData("Payless_GUI_Rentals", k, 28);
				 String COMM			= rwe.getCellData("Payless_GUI_Rentals", k, 29);
				 String COYY			= rwe.getCellData("Payless_GUI_Rentals", k, 30);
				 String MOPREASON		= rwe.getCellData("Payless_GUI_Rentals", k, 31);
				 String RETURNSTATION	= rwe.getCellData("Payless_GUI_Rentals", k, 32);
				 String RETURNDATE		= rwe.getCellData("Payless_GUI_Rentals", k, 33);
				 String RETURNTIME		= rwe.getCellData("Payless_GUI_Rentals", k, 34);
				 String CARGROUP		= rwe.getCellData("Payless_GUI_Rentals", k, 35);
				 String AWD				= rwe.getCellData("Payless_GUI_Rentals", k, 36);
				 String COUPON			= rwe.getCellData("Payless_GUI_Rentals", k, 37);
				 String FTNTYPE			= rwe.getCellData("Payless_GUI_Rentals", k, 38);
				 String FTNUMBER		= rwe.getCellData("Payless_GUI_Rentals", k, 39);
				 String REMARKS			= rwe.getCellData("Payless_GUI_Rentals", k, 40);
				 String COMVA			= rwe.getCellData("Payless_GUI_Rentals", k, 41);
				 String COMILEAGE		= rwe.getCellData("Payless_GUI_Rentals", k, 42);
				 String INSURANCE		= rwe.getCellData("Payless_GUI_Rentals", k, 43);
				 String COUNTERPRODUCT	= rwe.getCellData("Payless_GUI_Rentals", k, 44);
				 String ADRLASTNAME		= rwe.getCellData("Payless_GUI_Rentals", k, 45);
				 String ADRFIRSTNAME	= rwe.getCellData("Payless_GUI_Rentals", k, 46);
				 String ADRDOB			= rwe.getCellData("Payless_GUI_Rentals", k, 47);
				 String ADRCOUNTRY		= rwe.getCellData("Payless_GUI_Rentals", k, 48);
				 String ADRSTATE		= rwe.getCellData("Payless_GUI_Rentals", k, 49);
				 String ADRDRLICNO		= rwe.getCellData("Payless_GUI_Rentals", k, 50);
				 String ADREXPMM		= rwe.getCellData("Payless_GUI_Rentals", k, 51);
				 String ADREXPYY		= rwe.getCellData("Payless_GUI_Rentals", k, 52);
				 String ADRADDR1		= rwe.getCellData("Payless_GUI_Rentals", k, 53);
				 String ADRADDR2		= rwe.getCellData("Payless_GUI_Rentals", k, 54);
				 String ADRADDR3		= rwe.getCellData("Payless_GUI_Rentals", k, 55);
				 String ADRTELEPHONE1	= rwe.getCellData("Payless_GUI_Rentals", k, 56);
				 String ADRTELEPHONE2   = rwe.getCellData("Payless_GUI_Rentals", k, 57);
				 String ADRCCDC			= rwe.getCellData("Payless_GUI_Rentals", k, 58);
				 String ADRCARDNO		= rwe.getCellData("Payless_GUI_Rentals", k, 59);
				 String ADRCCEXPMM		= rwe.getCellData("Payless_GUI_Rentals", k, 60);
				 String ADRCCEXPYY		= rwe.getCellData("Payless_GUI_Rentals", k, 61);

				 
			if (Execute.equals("Y"))
			{
				
				
				String ScreenshotPath = prop.getProperty("ScreenshotPayless");	
				
				//* Open GUI URL's */
					// System.out.println(" token URL value : " + tokenURL);
					//*******Screenshot path and test name*********//
					
					String testcasename = TCName;
					String xfilepath = prop.getProperty("ExcelPathRentalPayless") +testcasename+ ".xlsx";
					test = extent.createTest(TCName);
						
					functions.openURL(thinClient);
					/* Login */
					//functions.login(uName, pswd);
					functions.navigateToTab("ReservationRates");
					Thread.sleep(2000);	
					
				 driver.findElement(By.xpath("//a[@data-target='#menulist\\:checkoutlink']")).click();
				 System.out.println("Clicked on Checkout Button");
				 Thread.sleep(3000);
				 System.out.println(driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).isDisplayed());
				 driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).click();
				 driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).sendKeys(ResNo);
				 System.out.println("Entered the existing Reservation");
				 driver.findElement(By.xpath("//input[@id='searchCommandLink']")).click();
				 System.out.println("Clicked on the Search Button");
				 functions.ScreenCapturedate(ScreenshotPath,TCName);
				 Thread.sleep(20000);
				 System.out.println( driver.findElement(By.xpath("//div[@id='generalInfo']//div[@class='modal-content']//div[1]//span[1]")).isDisplayed());
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//div[@id='generalInfo']//div[@class='modal-content']//div[1]//span[1]")).click();
				 Thread.sleep(2000);
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).click();
				  String Country = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).getAttribute("value");
				  System.out.println("The country is " + Country);
				  if (Country.isEmpty())
				  {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys(COCountry);
				 System.out.println("Entered the Country");
				  }
				  
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).click();
				 String State =driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).getAttribute("value");
				 System.out.println("The state is " +State );
				 if (State.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys(COState);
				 System.out.println("Entered the State");
				 }
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).click();
				 String Licence = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).getAttribute("value");
				 System.out.println("The Licence entered id " +Licence);
				 if (Licence.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys(CODRLICNO);
				 System.out.println("Entered the Licence Number");
				 }
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).click();
				 String DOB= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).getAttribute("value");
				 System.out.println("The Date od Birth entered is "+DOB);
				 if (DOB.isEmpty())
				 {	 
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys(CODOB);
				 System.out.println("Entered the Date of Birth");
				 }
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).click();
				 String Company = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).getAttribute("value");
				 System.out.println("The Company entered is "+Company);
				 if (Company.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys(COCOMPANY);
				 System.out.println("Entered the Company");
				 }
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).click();
				 String Add1 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).getAttribute("value");
				 if (Add1.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys(COADDR1);
				 System.out.println("Entered the Address 1");
				 }
				 String Add2= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).getAttribute("value");
				 if (Add2.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys(COADDR2);
				 System.out.println("Entered the Address 2");
				 }
				 String Add3 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).getAttribute("value");
				 if (Add3.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys(COADDR3);
				 System.out.println("Entered the Address 3");
				 }
				 String Contact =driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).getAttribute("value");
				 if (Contact.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys(COCONTACTINFO);
				 System.out.println("Entered the Contact Info");
				 }
				 if(driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).isEnabled())
				 {
				 String Email= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).getAttribute("value");
				 if (Email.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys(COEMAIL);
				 System.out.println("Entered the Email");
				 }
				 }
				 if(FTNTYPE != NULL)
				 {
				 String ftntp= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).getAttribute("value");
				 if (ftntp.isEmpty())
				 { 
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).sendKeys(FTNTYPE);
				 System.out.println("Entered the FTN Type");
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftNumber']")).sendKeys(FTNUMBER);
				 System.out.println("Entered the FTN Number");
				 }
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:verifiedFTN']")).click();
				 System.out.println("Verified the FTN number");
				 }
				 
				// driver.findElement(By.xpath("//input[@id='glyphicon glyphicon-plus']")).click();
				 
				 Select s= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:payMethod")));
				 s.selectByValue(COMOP);
				 System.out.println("Entered the Method of Payment");
				 Select s2= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType")));
				 s2.selectByValue(COCCDC);
				 System.out.println("Entered the Type of Payment");
				 String CardNumber = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).getAttribute("value");
				 if (CardNumber.isEmpty())
				 {
					
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).sendKeys(CARDNO);
				 System.out.println("Entered the Card Number");
				 }
				 String Month = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).getAttribute("value");
				 if (Month.isEmpty())
				 {
				 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).sendKeys(COMM);
					System.out.println("Entered the Expiry Month");
				 }
				 String year= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).getAttribute("value");
				 if (year.isEmpty())
				 {
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).sendKeys(COYY);
					System.out.println("Entered the Expiry Year");
				 }
				 
					Select s3= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason")));
					s3.selectByValue(MOPREASON);
					System.out.println("Entered the Reason for payment");
					//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:mopCCI']")).sendKeys("1234");
					//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:authorization']")).sendKeys("OK/1234");
					//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).clear();
					//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).sendKeys("GB/C");
					String RateCode = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).getAttribute("value");
					rwe.setCellData("Payless_GUI_Rentals", k, 62,RateCode);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(COMVA);
					System.out.println("Entered the MVA number");
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(COMILEAGE);
					System.out.println("Entered the Mileage");
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					if(driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).isDisplayed())
					{
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).click();
					}
					Thread.sleep(7000);
					if(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).isDisplayed())
					{
					String OptMSG= driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getAttribute("value");
					System.out.println("The Output message in vechicle continue screen is "+ OptMSG);
					if(OptMSG.equals("NET RATE ERROR")) 
					{
					System.out.println("ERROR MESSAGE "+OptMSG+" is Displayed");
					rwe.setCellData("Payless_GUI_Rentals", k, 70, OptMSG);
					rwe.setCellData("Payless_GUI_Rentals", k, 71, "FAIL");
					test.log(Status.FAIL, "Fail");
					
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					System.out.println("The Screenshot is taken");
					driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton")).click();
					driver.close();
					break;
					}
					else
						if (OptMSG.equals("START DATE INVALID FOR RATE"))
						{
						System.out.println("ERROR MESSAGE "+OptMSG+" is Displayed");
						rwe.setCellData("Payless_GUI_Rentals", k, 70, OptMSG);
						rwe.setCellData("Payless_GUI_Rentals", k, 71, "FAIL");
						test.log(Status.FAIL, "Fail");
						
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						System.out.println("The Screenshot is taken");
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton")).click();
						driver.close();
						break;	
							
						}
					}
					
				
					Thread.sleep(20000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					if (driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).isDisplayed())
					{
 				  	driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
 				  	System.out.println("Clicked on Continue Button");
					Thread.sleep(20000);
					}
					Thread.sleep(20000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					if(driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).isDisplayed())
					{
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).click();
					System.out.println("Clicked on Complete Checkout Button");
					Thread.sleep(7000);
					}
					
					Thread.sleep(7000);
					try
					{
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						 if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())
								 
						 {
							String str1= driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
							System.out.println("The message displayed is "+str1);
							if (str1.contains("Please enter the Credit ID security code (CVV/CCV)"))
							{
								if(COCCDC.equals("CA"))
					               {
					                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("1234");
					   
					                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
					                    Thread.sleep(12000);
					                    
					               }
					               else
					               {
					                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("123");
					                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
					                    Thread.sleep(12000);
					                    
					               }
								Thread.sleep(5000);
								 if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())								                                             
								 {
									String str2=driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
									System.out.println("The message displayed is "+str2); 
									if(str2.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION"))
										              
											{
										driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
										 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
										 driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
										 Thread.sleep(12000);
											}
								 }
								
								
							}
							else
								functions.ScreenCapturedate(ScreenshotPath,TCName);
								if (str1.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION"))
							{
								driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
								 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
								 driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
								 Thread.sleep(6000);
							}
					}
						 }
					catch (Exception e)
					{
						e.printStackTrace();
					}
				 
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//download the rental pdf
					Thread.sleep(2000);
					driver.findElement(By.xpath("//span[@class='ui-button-text ui-c'][contains(text(),'Close')]")).click();
					
					      
				 // Checkout complete Screen and update in Excel 
					   
					  //tring screenname = driver.findElement(By.xpath("//h5[@class='completeCheckout ng-binding']")).getText();
							   
						Thread.sleep(7000);			                    
					   //String strCOLNFNGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteName")).getText();
					   String strCOLNFNGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 63, strCOLNFNGetText);
		               System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
		              // String strCOVehicleModelGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle")).getText();
		               String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 64, strCOVehicleModelGetText);
		               System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
		               //String strCOResNoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum")).getText();
		               String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 65, strCOResNoGetText);
		               System.out.println(" Reservation No value is " + strCOResNoGetText);
		               //String strCOMVAGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA")).getText();
		               String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 66, strCOMVAGetText);
		               System.out.println(" MVA No value is " + strCOMVAGetText);
		               //String strCORANoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber")).getText();
		               String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 67, strCORANoGetText);
		               System.out.println(" RA Number value is " + strCORANoGetText);
		              // String strCOLicensePlateNoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate")).getText();
		               String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 68, strCOLicensePlateNoGetText);
		               System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
		               //String strCOEstimatedCostGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost")).getText();
		               String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 69, strCOEstimatedCostGetText);
		               System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
		               //String strCOSystemMsgGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteOut")).getText();
		               String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
		               Thread.sleep(7000);
		               rwe.setCellData("Payless_GUI_Rentals", k, 70, strCOSystemMsgGetText);
		               System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
		               Thread.sleep(7000);
		               
					
					
		               functions.ScreenCapturedate(ScreenshotPath,TCName);
					System.out.println("The Screenshot is taken");	
					Thread.sleep(5000);
					if (driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).isDisplayed())
		               {
						driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
		               System.out.println("Clicked on Done Button");
		               Thread.sleep(7000);
		               }
					//String ScreenShotPath2 = "C:\\Selenium\\Screenshots\\Avis\\Avis_CreateRentals\\";
					//functions.ScreenCapturedate(ScreenshotPath,TCName);
					System.out.println("The Screenshot is taken");
					//driver.findElement(By.xpath("//input[@id='headerLogOutButton']")).click();
					//driver.findElement(By.xpath("//button[@id='logoutForm:yesLogoutButton']")).click();
					//driver.close();
					Thread.sleep(2000);
					//functions.closeWindows();
					//test = extent.createTest(TCName);	
                    if(rwe.getCellData("Payless_GUI_Rentals", k, 64).isEmpty())
                    {
						rwe.setCellData("Payless_GUI_Rentals", k, 71,"FAIL");
				test.log(Status.FAIL, "Fail");
					} else {
						rwe.setCellData("Payless_GUI_Rentals", k, 71, "PASS");
				test.log(Status.PASS, "Pass");
					}
                  //Capturing all the screenshots in excel sheet
					FileOutputStream fileOut = null;
					int cntr =0;
					int row = 0;
					try {

					       Workbook wb = new XSSFWorkbook();
					      Sheet sheet = wb.createSheet("Ouput");
					       // FileInputStream obtains input bytes from the image file
					       String[] pathnames;

					       // Creates a new File instance by converting the given pathname string
					       // into an abstract pathname
					       File f = new File(prop.getProperty("ScreenshotPayless"));

					       // Populates the array with names of files and directories
					       pathnames = f.list();

					       // For each pathname in the pathnames array
					       for (String pathname : pathnames) {
					             // Print the names of files and directories

					             InputStream inputStream = new FileInputStream(
					            		 prop.getProperty("ScreenshotPayless")+pathname);
					             byte[] bytes = IOUtils.toByteArray(inputStream);
					             
					             // Adds a picture to the workbook
					       
					             int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
					             
					             // close the input stream
					             
					             // Returns an object that handles instantiating concrete classes
					             CreationHelper helper = wb.getCreationHelper();
					             // Creates the top-level drawing patriarch.
					             Drawing drawing = sheet.createDrawingPatriarch();

					             // Create an anchor that is attached to the worksheet
					             ClientAnchor anchor = helper.createClientAnchor();

					             // create an anchor with upper left cell _and_ bottom right cell
					             System.out.println(cntr);
					             anchor.setCol1(cntr); // Column B
					             anchor.setRow1(cntr+row ); // Row 3
					           

					             // Creates a picture
					             Picture pict = drawing.createPicture(anchor, pictureIdx);

					             // Reset the image to the original size
					             pict.resize(); //don't do that. Let the anchor resize the image!

					             // Create the Cell B
					             Cell cell = sheet.createRow(cntr).createCell(cntr);

					             // Write the Excel file

					             
					             fileOut = new FileOutputStream(xfilepath);
					             wb.write(fileOut);
					             row = row +40 ;
					             
					             inputStream.close();
					       }
					       fileOut.close();

					} catch (IOException ioex) {
					       System.out.println(ioex);
					}
			 
			}//end of if statement 
	

		else {
			System.out.println("Execution status is N for iteration " + k + "...");
		}
			    
	}//end of for statement


	functions.logout();
	
	Thread.sleep(1000);
	functions.closeWindows();
} finally{
extent.flush();
}
}
}

	
	