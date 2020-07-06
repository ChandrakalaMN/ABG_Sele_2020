package checkIn_Delayed;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;

import manifest.Extentmanager;

import pomClasses.GUIFunctions;
import pomClasses.ReadWriteExcel;

/**
'#################################################################################################################################
'## SCRIPT NAME:                           AVIS_GUI_CheckIn with Split Billing
'## BRAND:                                         AVIS
'## DESCRIPTION:                           Perform a Delayed CheckIn of Rental in different regions for Payless.
'## FUNCTIONAL AREA :                      CheckIn Screen
'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
'## OUTPUT:                                Rental should be created successfully.

'##################################################################################################################################
 **/


public class Avis_checkIn_Delay {
	
		ExtentReports extent;
		ExtentTest test;
	
		@BeforeTest
		public void startReport()
		{
			extent = Extentmanager.GetExtent();
			//test = extent.createTest("GUI");
		}
	
	//	public static void main(String[] args) throws Throwable {
		@Test
		public void test() throws Exception
		{
			try
			{
				
			   for (int K=1; K<=2; K++)
			   {
			
					ReadWriteExcel rwe 		= new ReadWriteExcel("C:\\GUI_Testing\\TestData\\DelayedCheckIn\\AVIS_GUITestData_DelayedCheckIn.xlsx");
					System.setProperty("webdriver.chrome.driver", "C:\\GUI_Testing\\WebDriver\\chromedriver.exe");
					WebDriver driver 		= new ChromeDriver();
					GUIFunctions functions 	= new GUIFunctions(driver);
					String ScreenShotPath 	= "C:\\GUI_Testing\\ScreenShot\\Avis\\DelayedCheckIn\\";
					
					String Execute 			= rwe.getCellData("DelayedCheckIn_Avis", K,  1);
					String ENVIRONMENT 		= rwe.getCellData("DelayedCheckIn_Avis", K,  2);
					String TESTCASENAME		= rwe.getCellData("DelayedCheckIn_Avis", K,  3);
					String THINCLIENTURL	= rwe.getCellData("DelayedCheckIn_Avis", K,  5);
					String Region			= rwe.getCellData("DelayedCheckIn_Avis", K,  6);
					String USERID 			= rwe.getCellData("DelayedCheckIn_Avis", K,  7);
					String PASSWORD 		= rwe.getCellData("DelayedCheckIn_Avis", K,  8);
					String RANO				= rwe.getCellData("DelayedCheckIn_Avis", K,  9);
					String RETURNDATE		= rwe.getCellData("DelayedCheckIn_Avis", K, 10);
					String RETURNTIME		= rwe.getCellData("DelayedCheckIn_Avis", K, 11);
					
					if (Execute.equals("N"))
					{
						driver.close();
					}
					
					if (Execute.equals("Y"))
					{
						driver.get(THINCLIENTURL + Region);
						driver.findElement(By.name("username")).sendKeys(USERID);
						driver.findElement(By.name("PASSWORD")).sendKeys(PASSWORD);
						
						driver.findElement(By.cssSelector("body > table:nth-child(7) > tbody > tr > td:nth-child(2) > form > table > tbody > tr:nth-child(3) > td > input[type=\"SUBMIT\"]:nth-child(3)")).click();
						driver.navigate().refresh();
						Thread.sleep(9000);
						
						driver.manage().window().maximize();
						Thread.sleep(2000);
						
						driver.findElement(By.xpath("//*[@id='menubar']/ul/li[2]/a")).click();
						Thread.sleep(2000);
						
						System.out.println(ENVIRONMENT);
						System.out.println("Ra number is"+RANO);
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:raNumber']")).sendKeys(RANO);
						Thread.sleep(20000);
						
						driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
						Thread.sleep(2000);
						
						WebElement MlgOut = driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:mileageOut']"));
						String MlgOutVal = MlgOut.getAttribute("value");
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:mileage']")).sendKeys(MlgOutVal);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:fuelPurchase']")).sendKeys("Yes");
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:returnDate1_hid']")).sendKeys(RETURNDATE);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:inTime1']")).sendKeys(RETURNTIME);
						
					// Click on to the Complete CheckIn
						driver.findElement(By.xpath("//*[@id='footerForm:completeCheckIn']")).click();
						Thread.sleep(10000);
						
						WebElement CashChekIn = driver.findElement(By.xpath("//*[@id='checkinRepromptDialog:checkInCashRepromptForm:checkInCashRepromptPanel']/div/div[1]/h5"));
						if (CashChekIn.isDisplayed())
						{
						    String getCashChekInText = CashChekIn.getText();
						    System.out.println(" CashChekIn Message " + getCashChekInText);

						    if (getCashChekInText.contains("Cash Check In"))
						    {
						    	WebElement Cash = driver.findElement(By.id("checkinRepromptDialog:checkInCashRepromptForm:checkInAmtDue"));
						    	String CashValue = Cash.getText();
						    	driver.findElement(By.id("checkinRepromptDialog:checkInCashRepromptForm:checkInCashReceived")).clear();
						    	driver.findElement(By.id("checkinRepromptDialog:checkInCashRepromptForm:checkInCashReceived")).sendKeys(CashValue);
								driver.findElement(By.id("checkinRepromptDialog:checkInCashRepromptForm:acceptCheckInCashReceived")).click();
								Thread.sleep(3000);
						    }
						}
						
						WebElement MOP = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
						if (MOP.isDisplayed())
						{
						    String getMOPText = MOP.getText();
						    System.out.println(" CashChekIn Message " + getMOPText);
						    
						    if (getMOPText.contains("PLEASE ENTER METHOD OF PAYMENT FOR THE UNDERPAYMENT OF"))
						    {
								driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
								Thread.sleep(10000);
						    }
						}
						
						WebElement Counters = driver.findElement(By.id("checkinRepromptDialog:checkInProductRepromptForm:repromptProductOut"));
						if (Counters.isDisplayed())
						{
						    String getCounterstext = Counters.getText();
						    System.out.println(" Get text value is: " + getCounterstext);
						    Thread.sleep(10000);
							    if (getCounterstext.contains("WERE ALL COUNTER PRODUCTS RETURNED AND UNDAMAGED?"))
							    {
									driver.findElement(By.id("checkinRepromptDialog:checkInProductRepromptForm:cpdYesBtn")).click();
									Thread.sleep(15000);
							    }
						}
						
						if (Counters.isDisplayed())
						{
						 			driver.findElement(By.id("checkinRepromptDialog:checkInProductRepromptForm:cpdYesBtn")).click();
									Thread.sleep(15000);
						}
						
						WebElement Valueis = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
						if (Valueis.isDisplayed())
						{
						    String getValueistext = Valueis.getText();
						    System.out.println(" Get text value is: " + getValueistext);
						    Thread.sleep(10000);
							    if (getValueistext.contains("C/I CAR INFO DISCREPANCIES"))
							    {
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(15000);
							    }
						}
						
						
						if (Valueis.isDisplayed())
						{
						    String getValueistext = Valueis.getText();
						    System.out.println(" Get text value is: " + getValueistext);
						    Thread.sleep(10000);
							    if (getValueistext.contains("UNDER PAYMENT"))
							    {
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(15000);
							    }
						}
						
//						WebElement ENTDATA = driver.findElement(By.xpath("//*[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']"));
//						String ent = ENTDATA.getText();
//						
//						if (ent.isEmpty()){
//							driver.findElement(By.xpath("//*[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).sendKeys("1234");
//							driver.findElement(By.xpath("//*[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
//							Thread.sleep(3000);
//						}
					
//							driver.findElement(By.xpath("//*[@id='printPreviewDialog:printPreviewForm:closePrintPreviewBtn']/span")).click();
//							Thread.sleep(3000);

//							Date d1= new Date();
//							SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
//							File src1= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
//							FileUtils.copyFile(src1,new File( ScreenShotPath +TESTCASENAME+ sdf1.format(d1)+"_Print"+".pgn"));
								
					//Complete Check In
						WebElement CompCheckIn = driver.findElement(By.xpath("//*[@id='checkInSuccessDlg:checkInSuccessForm:checkInSuccessPanel_content']/div[1]/h5"));
						if (CompCheckIn.isDisplayed());
						{
						    String getCheckInText = CompCheckIn.getText();
						    System.out.println(" Complated CheckIn " + getCheckInText);
						    Thread.sleep(5000);
						    if (getCheckInText.contains("Complete Check In"))
						    {
							
						    	String CCI = CompCheckIn.getText();
								rwe.setCellData("DelayedCheckIn_Avis", K, 12, CCI);
								
								Date d2= new Date();
								SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
								File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
								FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_CheckInCompleted"+".pgn"));
	
								
						    	driver.findElement(By.xpath("//*[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).click();
								Thread.sleep(3000);
						    }
					// Logout the application
					test = extent.createTest(TESTCASENAME);
					
					if (rwe.getCellData("DelayedCheckIn_Avis", K, 12).isEmpty())
					{
						rwe.setCellData("DelayedCheckIn_Avis", K, 13, "FAIL");
						test.log(Status.FAIL, "Fail");
					}
					else
					{
						test.log(Status.PASS, "Pass");
						rwe.setCellData("DelayedCheckIn_Avis", K, 13, "PASS");
					}

						functions.logout();
						driver.close();
						}	   
					}
			else
			{
			   System.out.println("Execution status is N for iteration "+K+"...");
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