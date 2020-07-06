package AVIS.TestScripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;

import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;

/**
* '#############################################################################################################################
* '## SCRIPT NAME: AVIS_GUI_Modify_Rentals 
* '## BRAND: AVIS 
* '## DESCRIPTION: Modify Rentals for all the editable fields 
* '## FUNCTIONAL AREA : Display Rental Screen 
* '## PRECONDITION:All the required Test Data should be available in Test Data Sheet. 
* '##OUTPUT: Rentals should be modified successfully.
* 
*HISTORY 12-SEP-2018 - GUIFunctions class created for GUI Common
* functionalities and CR functionality
* '#############################################################################################################################
**/

public class AVIS_GUI_Modify_Rentals
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
	public void startReport()
	{
		extent = Extentmanager.GetExtent();
		//test = extent.createTest("GUI");
	}
	
	@SuppressWarnings("unlikely-arg-type")
	//public static void main(String[] args) throws IOException, Exception, FileNotFoundException {
	
	@Test
	public void test() throws Exception
	{
		try
		{
					
		// Read input from excel
		
		for (int k = 1; k <= 100; k++)
		{
			AVIS_GUI_Modify_Rentals avis = new AVIS_GUI_Modify_Rentals();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Users\\cmn\\Desktop\\Selenium\\TestData\\AVIS_GUITestData_DisplayModifyRental.xlsx");
			String Execute = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 2);
			
			if (Execute.equals("Y"))
			{
				System.setProperty("webdriver.chrome.driver","C:\\Users\\cmn\\Desktop\\Selenium\\chromedriver.exe");
				ChromeDriver driver = new ChromeDriver();
				GUIFunctions functions = new GUIFunctions(driver);
				int a = 27;
				//ChromeDriver driver = new ChromeDriver();
				//GUIWebDriverFunctions wdfunctions = new GUIWebDriverFunctions(driver);
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				System.out.println(" iteration " + k);
				String testcasename    = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 4);
				String tokenURL    	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 6);
				String clientURL       = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 7);
				String outSTA          = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 8);
				String thinClient      = clientURL+outSTA;
				String uName           = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 9);
				String pswd            = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 10);
				String Rano            = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 11);
				String Custcoun		   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 12);
				String Custstate	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 13);
				String CustLicno	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 14);
				String CustDOB		   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 15);
				String Custcompy	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 16);
				String CustAddr1	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 17);
				String CustAddr2	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 18);
				String CustAddr3	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 19);
				String Custcontno	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 20);
				String Custemail	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 21);
				String CustFtntype	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 22);
				String CustFtnno	   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 23);
				String Vehiclemilageout	= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 24);
				String Vehicleinfospace	= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 25);
				String Vehicleinfofuelsvc= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 26);
				String Vehicleinfofuellvl= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 27);
				String Vehicleinfodamaged= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 28);
				String Vehicleinfoaccrept= rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 29);
				String Reninfocodte		 = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 30);
				String Reninfocotme      = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 31);
				String Reninfoawd        = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 32);
				String Reninfocpn	     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 33);
				String Reninfortnstation = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 34);
				String Reninfortndate    = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 35);
				String Reninfortntime    = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 36);
				String Reninfodelivery   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 37);
				String Reninforatecode   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 38);
				String Reninforemarks    = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 39);
				String Reninfonewayfee   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 40);
				String Reninforemchkbx   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 41);
				String Reninfocomissio   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 42);
				String Reninfotax        = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 43);
				String Reninfodiscount   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 44);
				String Reninfomancalamt   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 45);
				String Reninfomancalamtdrpdwn   = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 46);
				String Paymentinfomop     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 47);
				String Paymentinfoccdc     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 48);
				String Paymentinfocardno     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 49);
				String Paymentinfomm     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 50);
				String Paymentinfoyy     = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 51);
				String Paymentinfomoprsn    = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 52);
				String AdjCustsercer = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 53);
				String AdjCustmoneyoff = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 54);
				String AdjCustbreakdown = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 55);
				String AdjCustseccpn = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 56);
				String AdjCustdamproperty = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 57);
				String AdjCustfuel = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 58);
				String AdjCustfuelamount = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 59);
				String AdjCustothreason = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 60);
				String AdjCustothamount = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 61);
				String Misparking = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 62);
				String Mishandling = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 63);
				String Misaccrep = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 64);
				String Misluggage = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 65);
				String Misphoneacc = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 66);
				String Miswheeldam = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 67);
				String Miscleaningmain = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 68);
				String Misparkingfines = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 69);
				String Mispickupfees = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 70);
				String Mispickuptowing = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 71);
				String Misothers = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 72);
				String Misoneway = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 73);
				String ProtsCovagesInsur = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 74);
				String Counterprd = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 75);
				String Messagetype = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 76);
				String Messagebox = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 77);
				
				
				
				/* Open GUI URL's */
				driver.get(thinClient);
				
				/* Login */
				functions.login(uName, pswd);
				driver.navigate().refresh();
				functions.navigateToTab("DisplayRental");
				Thread.sleep(5000L);
				
				//Entering the RA number in search field
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:rentalSearchForm:searchString']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:rentalSearchForm:searchString']")).sendKeys(Rano);
				driver.findElement(By.xpath("//input[@id='searchCommandLinkDisplayRental']")).click();
				Thread.sleep(3000);
				
			
				
				//Modifying Address 1
				if(!CustAddr1.isEmpty())
				{
					Thread.sleep(1000);
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress1']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress1']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress1']")).sendKeys(CustAddr1);
					Thread.sleep(3000);
				}
				//Modifying Address 2
				else if(!CustAddr2.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress2']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress2']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress2']")).sendKeys(CustAddr2);
					Thread.sleep(3000);
				}
				//Modifying Address 3
				else if(!CustAddr3.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress3']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress3']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:maddress3']")).sendKeys(CustAddr3);
					Thread.sleep(3000);
				}
				
				//Modifying Contact number
				else if(!Custcontno.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcontactInfo']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcontactInfo']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcontactInfo']")).sendKeys(Custcontno);
					Thread.sleep(3000);
					
				}
				//Modifying Email address
				else if(!Custemail.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:wizconEmail']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:wizconEmail']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:wizconEmail']")).sendKeys(Custemail);
					Thread.sleep(3000);
				}
				
				//Modifying FTN type
				else if(!CustFtntype.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftnType']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftnType']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftnType']")).sendKeys(CustFtntype);
					Thread.sleep(3000);
				}
				
				//Modifying FTN number
				else if(!CustFtnno.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftNumber']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftNumber']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mftNumber']")).sendKeys(CustFtnno);
					Thread.sleep(3000);	
				}
				
				//Modifying Mileage out
				else if(!Vehiclemilageout.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mrentalPickupMileage']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mrentalPickupMileage']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mrentalPickupMileage']")).sendKeys(Vehiclemilageout);
					Thread.sleep(3000);
				}
				
				//Modifying Space number
				else if(!Vehicleinfospace.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mspaceNumber']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mspaceNumber']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mspaceNumber']")).sendKeys(Vehicleinfospace);
					Thread.sleep(3000);
				}
				
				//
				//Modifying Fuel SVC dropdown
				else if(!Vehicleinfofuelsvc.isEmpty())
				{
					Select s = new Select(driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mfuelSvc")));
					//s.selectByValue("Y");
					if(Vehicleinfofuelsvc.contains("Yes"))
					{
						s.selectByValue("Y");
						//If you get error message on selecting the fuel service option
						System.out.println(driver.switchTo().alert().getText());
						driver.switchTo().alert().accept();
					}
						else 
						{
							s.selectByValue("N");
							//If you get error message on selecting the fuel service option
							System.out.println(driver.switchTo().alert().getText());
							driver.switchTo().alert().accept();
						}
					}
			
				//Modifying Fuel out
				else if(!Vehicleinfofuellvl.isEmpty())
				{
					
					//functions.selectFuelout(Vehicleinfofuellvl);
					Select fuelout = new Select(driver.findElement(By.xpath("//select[@id='menulist:dispormodContainer:dispormodForm:mfuelOut']")));
				    fuelout.selectByValue(Vehicleinfofuellvl);	
					Thread.sleep(3000);
				}
				
				//Modifying Vehicle damaged dropdown field
				else if(!Vehicleinfodamaged.isEmpty())
				{
					Select infodam=new Select(driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mvehicleDamaged")));
					if(Vehicleinfodamaged.contains("No"))
					{
						infodam.selectByValue("Y");
						Thread.sleep(2000);
					}
					else
					{
						infodam.selectByValue("N");
						Thread.sleep(2000);
					}
		
					
				}
				
				//Modifying Accident report dropdown field
				else if(!Vehicleinfoaccrept.isEmpty())
				{
					Select infoaccrept=new Select(driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:maccidentReported")));
					if(Vehicleinfoaccrept.contains("No"))
					{
						infoaccrept.selectByValue("Y");
						Thread.sleep(3000);
					}
					else
					{
						infoaccrept.selectByValue("N");
						Thread.sleep(2000);
					}
				}
				
				//Modifying checkin location and one way fee charges
				else if(!Reninfortnstation.isEmpty())
				{
					Thread.sleep(3000);
					System.out.println(driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInStation']")).isDisplayed());
					
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInStation']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInStation']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInStation']")).sendKeys(Reninfortnstation);
					Thread.sleep(3000);
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).sendKeys(Reninfonewayfee);
					Thread.sleep(3000);	
					driver.findElement(By.xpath("//input[@id='footerForm:saveModifyRentalStep1']")).click();
					Thread.sleep(7000);
					//Prompt appeared for one way fee
					String promptforoneway = driver.findElement(By.id("rentalRepromptDialog:repromptForm:repromptOut")).getText();
					if(promptforoneway.contains("PRESS ENTER TO CONTINUE"))
					{
						driver.findElement(By.xpath("//input[@id='rentalRepromptDialog:repromptForm:rentalModRepromptCancelButton']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneway']")).sendKeys("750");
						
					}
				}
				
				//Modifying checkin date and time
				if(!Reninfortndate.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckinDate_hid']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckinDate_hid']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckinDate_hid']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckinDate_hid']")).sendKeys(Reninfortndate);
					Thread.sleep(3000);
				}
				
				//Modifying checkin time
				if(!Reninfortntime.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInTime']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInTime']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInTime']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mcheckInTime']")).sendKeys(Reninfortntime);
					Thread.sleep(3000);
				}
			
				//Modifying delivery dropdown box with option Preferred/FBO
				/**else if(!Reninfodelivery.isEmpty())
				{
					driver.findElement(By.xpath("//div[@id='menulist:dispormodlink']//div[@id='subreservationInfoPanel_content']//div[@class='form-group']//select[2]")).click();	
					Select delivery= new Select(driver.findElement(By.xpath("//div[@id='menulist:dispormodlink']//div[@id='subreservationInfoPanel_content']//div[@class='form-group']//select[2]")));
					if(Reninfodelivery.contains("Preferred"))
					{	
						delivery.selectByVisibleText("Preferred");
						Thread.sleep(3000);
					}
					
						else
						{
							delivery.selectByVisibleText("FBO");
						}
					
				}
				**/
				
				//Modifying Remarks field under Rental information section
				else if(!Reninforemarks.isEmpty())
				{
					driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mremarks")).isDisplayed();
					driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mremarks")).clear();
					driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mremarks")).click();
					driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mremarks")).sendKeys(Reninforemarks);
					Thread.sleep(3000);
					
				}
			
				//Modifying remove checkbox under Rental information section
				else if(!Reninforemchkbx.isEmpty())
				{
					//List<WebElement> els = driver.findElements(By.xpath("//input[@type='checkbox']"));
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:removeOneWay']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:removeOneWay']")).click();
					Thread.sleep(3000);
					
				}
				//Modifying commission checkbox
				else if(!Reninfocomissio.isEmpty())
				{
					driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mcommission")).isDisplayed();
					Select comission=new Select(driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:mcommission")));
					if(Reninfocomissio.contains("No"))
					{
						comission.selectByValue("N");
						Thread.sleep(3000);
					}
					else
					{
						comission.selectByValue("Y");
						Thread.sleep(3000);
					}
				}
				
				//Modifying Tax field under rental information
				else if(!Reninfotax.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mtax']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mtax']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mtax']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:mtax']")).sendKeys(Reninfotax);
					Thread.sleep(3000);
				}
				
				//Modifying Payment information
				else if(!Paymentinfomop.isEmpty())
				{
					//selecting credit card
					Select payment= new Select(driver.findElement(By.xpath("//select[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccType']")));
					payment.selectByVisibleText(Paymentinfoccdc);
					Thread.sleep(3000);
					
					//entering credit card number
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccNumber']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccNumber']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccNumber']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccNumber']")).sendKeys(Paymentinfocardno);
					Thread.sleep(3000);
					
					//Entering Month
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireMonth']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireMonth']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireMonth']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireMonth']")).sendKeys(Paymentinfomm);
					Thread.sleep(3000);
					
					//Entering Year
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireYear']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireYear']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireYear']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:ccExpireYear']")).sendKeys(Paymentinfoyy);
					Thread.sleep(3000);
					
					//Selecting Reason
					Select reason= new Select(driver.findElement(By.xpath("//select[@id='menulist:dispormodContainer:dispormodForm:creditCard:swipe:paymentReason']")));
					reason.selectByVisibleText(Paymentinfomoprsn);
					Thread.sleep(3000);
					
				}
				//Modifying Adjustment section
				
				if(!AdjCustsercer.isEmpty())
				{
				//Clicking on adjustment expand button
				driver.findElement(By.xpath("//div[@id='madjustmentsPanel_header']/span[2]/span")).isDisplayed();
				driver.findElement(By.xpath("//div[@id='madjustmentsPanel_header']/span[2]/span")).click();
					
				//entering cust service certificate
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:custServiceCertificate']")).isDisplayed();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:custServiceCertificate']")).isEnabled();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:custServiceCertificate']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:custServiceCertificate']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:custServiceCertificate']")).sendKeys(AdjCustsercer);
				Thread.sleep(3000);
				
				//Entering customer money off
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneyOff']")).isDisplayed();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneyOff']")).isEnabled();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneyOff']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneyOff']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:moneyOff']")).sendKeys(AdjCustmoneyoff);
				Thread.sleep(3000);
				
				//Entering breakdown
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:breakDownComp']")).isDisplayed();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:breakDownComp']")).isEnabled();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:breakDownComp']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:breakDownComp']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:breakDownComp']")).sendKeys(AdjCustbreakdown);
				Thread.sleep(3000);
				
				//Entering Coupon
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:adjCoupon']")).isDisplayed();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:adjCoupon']")).isEnabled();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:adjCoupon']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:adjCoupon']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:adjCoupon']")).sendKeys(AdjCustseccpn);
				Thread.sleep(3000);
				
				//Damaged Property
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:damagedProperty']")).isDisplayed();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:damagedProperty']")).isEnabled();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:damagedProperty']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:damagedProperty']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:damagedProperty']")).sendKeys(AdjCustdamproperty);
				Thread.sleep(3000);
			
				}
				
				//Modifying Miscellaneous Charges section
				if(!Misparking.isEmpty())
				{
					//Clicking on Miscellaneous Charges expand button
					driver.findElement(By.xpath("//div[@id='mMiscChargesPanel_header']/span[2]/span")).isDisplayed();
					driver.findElement(By.xpath("//div[@id='mMiscChargesPanel_header']/span[2]/span")).click();	
					Thread.sleep(3000);
					
					//parking garage storage
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingGarageStorage']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingGarageStorage']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingGarageStorage']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingGarageStorage']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingGarageStorage']")).sendKeys(Misparking);
					Thread.sleep(3000);
					
					//Handling Fee Cash Qual 
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:handlingFeeCash']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:handlingFeeCash']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:handlingFeeCash']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:handlingFeeCash']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:handlingFeeCash']")).sendKeys(Mishandling);
					Thread.sleep(3000);
					
					//Accident Repairs
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:accidentRepairs']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:accidentRepairs']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:accidentRepairs']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:accidentRepairs']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:accidentRepairs']")).sendKeys(Misaccrep);
					Thread.sleep(3000);
					
					//Luggage Rack
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:skiLuggageRacks']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:skiLuggageRacks']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:skiLuggageRacks']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:skiLuggageRacks']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:skiLuggageRacks']")).sendKeys(Misluggage);
					Thread.sleep(3000);
					
					//Phone Accessories  
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscChains']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscChains']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscChains']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscChains']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscChains']")).sendKeys(Misphoneacc);
					Thread.sleep(3000);
					
					//Wheel damage
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:tireWheelDamage']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:tireWheelDamage']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:tireWheelDamage']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:tireWheelDamage']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:tireWheelDamage']")).sendKeys(Miswheeldam);
					Thread.sleep(3000);
					
					//Cleaning/Maint
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:cleaningMaintenance']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:cleaningMaintenance']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:cleaningMaintenance']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:cleaningMaintenance']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:cleaningMaintenance']")).sendKeys(Miscleaningmain);
					Thread.sleep(3000);
					
					//Parking Fines/Citations
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingFinesCitations']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingFinesCitations']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingFinesCitations']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingFinesCitations']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:parkingFinesCitations']")).sendKeys(Misparkingfines);
					Thread.sleep(3000);
					
					//Delivery/Pickup fee
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:deliveryPickupFee']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:deliveryPickupFee']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:deliveryPickupFee']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:deliveryPickupFee']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:deliveryPickupFee']")).sendKeys(Mispickupfees);
					Thread.sleep(3000);
					
					//Towing
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscTowing']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscTowing']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscTowing']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscTowing']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscTowing']")).sendKeys(Mispickuptowing);
					Thread.sleep(3000);
					
					//Other
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOther']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOther']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOther']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOther']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOther']")).sendKeys(Misothers);
					Thread.sleep(3000);
					
					//One Way
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOneWay']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOneWay']")).isEnabled();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOneWay']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOneWay']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:dispormodContainer:dispormodForm:miscOneWay']")).sendKeys(Misoneway);
					Thread.sleep(3000);
				
				}
				
				//Modifying Insurance
				if (ProtsCovagesInsur.isEmpty()) {
					System.out.print("No Insurance selected");
				} else {
					String[] Insurancevalue = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 74).split("-");
					for (String e : Insurancevalue) {
						WebDriverWait wait1 = new WebDriverWait(driver, 10);
						if (e.equalsIgnoreCase("LDW")) {
							WebElement insurace1 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:lossDamageWaiver"));
							Select insLDW = new Select(insurace1);
							wait1.until(ExpectedConditions.visibilityOf(insurace1));
							if (insurace1.isDisplayed()) {
								insLDW.selectByVisibleText("Yes");
							} else {
								break;
							}

						} else if (e.equalsIgnoreCase("PAI")) {
							WebElement insurace2 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:personalAccidentInsurance"));
							Select insPAI = new Select(insurace2);
							wait1.until(ExpectedConditions.visibilityOf(insurace2));
							if (insurace2.isDisplayed()) {
								insPAI.selectByVisibleText("Yes");
							} else {
								break;
							}
						} else if (e.equalsIgnoreCase("PEP")) {
							WebElement insurace3 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:personalEffectsProtection"));
							Select insPEP = new Select(insurace3);
							wait1.until(ExpectedConditions.visibilityOf(insurace3));
							if (insurace3.isDisplayed()) {
								insPEP.selectByVisibleText("Yes");
							} else {
								break;
							}

						} else if (e.equalsIgnoreCase("ALI")) {
							WebElement insurace4 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:additionalLiabilityInsurance"));
							Select insALI = new Select(insurace4);
							wait1.until(ExpectedConditions.visibilityOf(insurace4));
							if (insurace4.isDisplayed()) {
								insALI.selectByVisibleText("Yes");
							} else {
								break;
							}
						}
						}
					}

			/**
			 //Modifying message and saving the same
				if(!Messagetype.isEmpty())
				{
					String[] Message = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 76).split("-");
					for (String e : Message) {
						WebDriverWait wait1 = new WebDriverWait(driver, 10);
						if (e.equalsIgnoreCase("Note to Self")) {
							WebElement message1 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:nMessageType"));
							Select note = new Select(message1);
							wait1.until(ExpectedConditions.visibilityOf(message1));
							if (message1.isDisplayed()) {
								note.selectByVisibleText("Note to Self");
							} else {
								break;
							}
						}
							else if (e.equalsIgnoreCase("Cancel Reservation")) {
								WebElement message2 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:nMessageType"));
								Select cancel = new Select(message2);
								wait1.until(ExpectedConditions.visibilityOf(message2));
								if (message2.isDisplayed()) {
									cancel.selectByVisibleText("Cancel Reservation");
								} else {
									break;
								}
							}
								else if (e.equalsIgnoreCase("Message to Company")) {
									WebElement message3 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:nMessageType"));
									Select company = new Select(message3);
									wait1.until(ExpectedConditions.visibilityOf(message3));
									if (message3.isDisplayed()) {
										company.selectByVisibleText("Message to Company");
									} else {
										break;
									}
							}
								else if (e.equalsIgnoreCase("Extension Request")) {
									WebElement message4 = driver.findElement(By.id("menulist:dispormodContainer:dispormodForm:nMessageType"));
									Select request = new Select(message4);
									wait1.until(ExpectedConditions.visibilityOf(message4));
									if (message4.isDisplayed()) {
										request.selectByVisibleText("Extension Request");
									} else {
										break;
									}
							}**/
				
				//Click on continue button/Save button/done button
				driver.findElement(By.xpath("//input[@id='footerForm:saveModifyRentalStep1']")).click();
				Thread.sleep(9000);
				//Prompt appeared for one way fee
				//String promptforoneway = driver.findElement(By.id("rentalRepromptDialog:repromptForm:repromptOut")).getText();
				//System.out.println(promptforoneway);
				//Prompt appeared for one way fee
				WebElement Onewayprompt = driver.findElement(By.xpath("//input[@id='rentalRepromptDialog:repromptForm:rentalModRepromptButton']"));
				Thread.sleep(5000);
				if(Onewayprompt.isDisplayed())
				{
					//Click on enter button on one way pop up screen
					driver.findElement(By.xpath("//input[@id='rentalRepromptDialog:repromptForm:rentalModRepromptButton']")).click();
					Thread.sleep(3000);
				}
				
				//Additional products
				if (Counterprd.isEmpty()) {
					System.out.print("No Counter product selected");
				} else {
					String[] CPvalue = rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 75).split("-");
					for (String e : CPvalue) {
						WebDriverWait wait1 = new WebDriverWait(driver, 10);
						if (e.equalsIgnoreCase("ADD")) {
							WebElement CP1 = driver.findElement(By.id("productQuantity990"));
							Select CPADD = new Select(CP1);
							wait1.until(ExpectedConditions.visibilityOf(CP1));
							if (CP1.isDisplayed()) {
								CPADD.selectByVisibleText("1");
							} else {
								break;
							}

						} else if (e.equalsIgnoreCase("CBS")) {
							WebElement CP2 = driver.findElement(By.id("productQuantity31"));
							Select CPCBS = new Select(CP2);
							wait1.until(ExpectedConditions.visibilityOf(CP2));
							if (CP2.isDisplayed()) {
								CPCBS.selectByVisibleText("1");
							} else {
								break;
							}
						} else if (e.equalsIgnoreCase("CIS")) {
							WebElement CP3 = driver.findElement(By.id("productQuantity32"));
							Select CPCIS = new Select(CP3);
							wait1.until(ExpectedConditions.visibilityOf(CP3));
							if (CP3.isDisplayed()) {
								CPCIS.selectByVisibleText("1");
							} else {
								break;
							}

						} else if (e.equalsIgnoreCase("CSS")) {
							WebElement CP4 = driver.findElement(By.id("productQuantityYesNo3"));
							Select CPCSS = new Select(CP4);
							wait1.until(ExpectedConditions.visibilityOf(CP4));
							if (CP4.isDisplayed()) {
								CPCSS.selectByVisibleText("Yes");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("CS")) {
							WebElement CP5 = driver.findElement(By.id("productQuantity34"));
							Select CPCS = new Select(CP5);
							wait1.until(ExpectedConditions.visibilityOf(CP5));
							if (CP5.isDisplayed()) {
								CPCS.selectByVisibleText("1");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("ESP")) {
							WebElement CP6 = driver.findElement(By.id("productQuantityYesNo6"));
							Select CPESP = new Select(CP6);
							wait1.until(ExpectedConditions.visibilityOf(CP6));
							if (CP6.isDisplayed()) {
								CPESP.selectByVisibleText("Y");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("GPS")) {
							WebElement CP7 = driver.findElement(By.id("productQuantity37"));
							Select CPGPS = new Select(CP7);
							wait1.until(ExpectedConditions.visibilityOf(CP7));
							if (CP7.isDisplayed()) {
								CPGPS.selectByVisibleText("1");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("SSP")) {
							WebElement CP8 = driver.findElement(By.id("productQuantityYesNo8"));
							Select CPSSP = new Select(CP8);
							wait1.until(ExpectedConditions.visibilityOf(CP8));
							if (CP8.isDisplayed()) {
								CPSSP.selectByVisibleText("Y");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("RSN")) {
							WebElement CP9 = driver.findElement(By.id("productQuantityYesNo9"));
							Select CPRSN = new Select(CP9);
							wait1.until(ExpectedConditions.visibilityOf(CP9));
							if (CP9.isDisplayed()) {
								CPRSN.selectByVisibleText("Y");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("SKI")) {
							WebElement CP10 = driver.findElement(By.id("productQuantityYesNo10"));
							Select CPSKI = new Select(CP10);
							wait1.until(ExpectedConditions.visibilityOf(CP10));
							if (CP10.isDisplayed()) {
								CPSKI.selectByVisibleText("Y");
							} else {
								break;
							}
						}
						else if (e.equalsIgnoreCase("SXR")) {
							WebElement CP11 = driver.findElement(By.id("productQuantityYesNo11"));
							Select CPSXR = new Select(CP11);
							wait1.until(ExpectedConditions.visibilityOf(CP11));
							if (CP11.isDisplayed()) {
								CPSXR.selectByVisibleText("Y");
							} else {
								break;
							}
						}
						
						}
					}

				
				//Button for save
				driver.findElement(By.xpath("//input[@id='footerForm:modifyRental']")).click();
				Thread.sleep(5000);
				//Button on prompt screen Done
				driver.findElement(By.xpath("//input[@id='modifyRentalDialog:successModDialogForm:rentalModCompDone']")).click();
				Thread.sleep(3000);
				
				/**
				 * //If Rental number not found
				String Rentalnotfound = driver.switchTo().alert().getText();
				if(Rentalnotfound.contains("Rental not found"))
				{
					//driver.switchTo().alert().accept();
					//driver.findElement(By.id("rentalsearch:rentalSearchPopup:rentalSearchCancel")).click();
					Thread.sleep(3000);
					functions.logout();
					functions.closeWindows();
				}
				**/
				
			
				
				
				
				//Quick View
				int a1 =78;
				WebElement rental = driver.findElement(By.xpath("//span[@class='qv-link qvDiRaNo']"));
                try {
                       if (rental.getText().isEmpty()) {
                              WebDriverWait wait1 = new WebDriverWait(driver, 20);
                              wait1.until(ExpectedConditions.visibilityOf(rental));
                       }
                } catch (Exception e2) {
                       e2.printStackTrace();
                }
                WebElement table = driver.findElement(By.id("QuickView-qvDisplayRental"));
                ArrayList<WebElement> rows = (ArrayList<WebElement>) table.findElements(By.tagName("tr"));
                for (int i = 1; i < rows.size(); i++) {
                       ArrayList<WebElement> cells = (ArrayList<WebElement>) rows.get(i).findElements(By.tagName("td"));
                       for (int j = 0; j < cells.size(); j++) {
                              String val = cells.get(j).getText();
                              if (val.isEmpty()) {
                                     break;
                              } else {
                                     if (j == 2) {
                                    	 if(a1<=99)
                                    	 {
                                            val = val.replaceAll("[*]", ""); // Remover '*' before rates
                                            rwe.setCellData("AVIS_GUI_Modify_Rentals", k, a1, val);
                                            a1++;
                                    	 }
                                     }
                              }
                       }
                }

			//div[@id='QuickView-qvDisplayRental']//div[@class='panel-body']
			
				//taking screenshot
				String ScreenShotPath = "C:\\Users\\cmn\\Desktop\\Selenium\\Screenshots\\AVIS_GUITestData_DisplayModifyRental\\";
				functions.ScreenCapture(ScreenShotPath , testcasename);

				

				/*
				 * Log out and close tabs
				 */

				functions.logout();
				Thread.sleep(1000);
				functions.closeWindows();
				test = extent.createTest(testcasename);
				if (rwe.getCellData("AVIS_GUI_Modify_Rentals", k, 79).isEmpty()) {
					
					rwe.setCellData("AVIS_GUI_Modify_Rentals", k, 100, "FAIL");
					test.log(Status.FAIL, "Fail");
				} else {
					rwe.setCellData("AVIS_GUI_Modify_Rentals", k, 100, "PASS");
					test.log(Status.PASS, "Pass");
				}
			}
			else
			{
				System.out.println("Execution status is N for iteration " + k + "...");
			}

					
					}
		} finally {
			extent.flush();
			// TODO: handle finally clause
		}
				}
				
			}