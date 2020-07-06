package checkIn_SB;

	import java.io.File;
	import java.text.SimpleDateFormat;
	import java.util.Date;
	import java.util.concurrent.TimeUnit;

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
	'## DESCRIPTION:                           Perform a CheckIn of Rental in different regions for AVIS.
	'## FUNCTIONAL AREA :                      CheckIn Screen
	'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
	'## OUTPUT:                                Rental should be created successfully.
	
	'##################################################################################################################################
	 **/	
	
	public class Avis_checkIn_SB {
	
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
				
			   for (int K=1; K<=5; K++)
			   {
				
					ReadWriteExcel rwe = new ReadWriteExcel("C:\\GUI_Testing\\TestData\\CheckIn\\AVIS_GUITestData_CheckIn.xlsx");
					System.setProperty("webdriver.chrome.driver", "C:\\GUI_Testing\\WebDriver\\chromedriver.exe");
					WebDriver driver = new ChromeDriver();
					GUIFunctions functions = new GUIFunctions(driver);
					driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
					String ScreenShotPath = "C:\\GUI_Testing\\ScreenShot\\Avis\\CheckIn\\";
					
					String Execute 			= rwe.getCellData("CheckIn_Avis", K,  1);
					String TESTCASENAME		= rwe.getCellData("CheckIn_Avis", K,  3);
					String THINCLIENTURL	= rwe.getCellData("CheckIn_Avis", K,  5);
					String Region			= rwe.getCellData("CheckIn_Avis", K,  6);
					String USERID 			= rwe.getCellData("CheckIn_Avis", K,  7);
					String PASSWORD 		= rwe.getCellData("CheckIn_Avis", K,  8);
					String RANO				= rwe.getCellData("CheckIn_Avis", K,  9);
					String MOP				= rwe.getCellData("CheckIn_Avis", K, 10);
					String CARDTYPE			= rwe.getCellData("CheckIn_Avis", K, 11);
					String CARDNO			= rwe.getCellData("CheckIn_Avis", K, 12);
					String EXPMONTH			= rwe.getCellData("CheckIn_Avis", K, 13);
					String EXPYEAR			= rwe.getCellData("CheckIn_Avis", K, 14);
					String CKECKINOUT_AUTH	= rwe.getCellData("CheckIn_Avis", K, 15);
					String ADDR1			= rwe.getCellData("CheckIn_Avis", K, 16);
					String ADDR3			= rwe.getCellData("CheckIn_Avis", K, 17);
					String SBTYPE			= rwe.getCellData("CheckIn_Avis", K, 18);
				
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
						
						driver.manage().window().maximize();
					
						driver.findElement(By.xpath("//*[@id='menubar']/ul/li[2]/a")).click();
						Thread.sleep(2000);
						System.out.println("Ra number is"+RANO);
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:raNumber']")).sendKeys(RANO);
						Thread.sleep(20000);
											
						WebElement MlgOut = driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:mileageOut']"));
						String MlgOutVal = MlgOut.getAttribute("value");
						System.out.println(MlgOutVal);
						int foo = Integer.parseInt(MlgOutVal);
						int TotMlgOut = foo + 200;
						String MilageTot = "" +TotMlgOut;
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:mileage']")).sendKeys(MilageTot);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkinForm:fuelPurchase']")).sendKeys("Yes");
						
						if (driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireMonth")).isDisplayed() && driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireMonth")).isEnabled());
						{
							String GetMonth = driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireMonth")).getAttribute("value");
							if (GetMonth.isEmpty())
							{
								driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireMonth")).sendKeys("12");
							}
							
						}
						
						if (driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireYear")).isDisplayed() && driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireYear")).isEnabled());
						{
							String GetYear = driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireYear")).getAttribute("value");
							if (GetYear.isEmpty())
							{
								driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:ccExpireYear")).sendKeys("21");
							}
						}
						
						if (driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:paymentReason")).isDisplayed() && driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:paymentReason")).isEnabled());
						{
							driver.findElement(By.id("menulist:checkinContainer:checkinForm:CVSwipe:paymentReason")).sendKeys("No reader/reader not working");
						}
					
					// SplitBilling Screen
						driver.findElement(By.id("footerForm:checkInSplitBillingButton")).click();
						Thread.sleep(8000);
	
						WebElement OneWayFee = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
						
						if (OneWayFee.isDisplayed())
						{
							System.out.println("Property Identified inside Oneway Loop");
						    String getOneWayFeeText = OneWayFee.getText();
						    System.out.println(" One way fee get text value " + getOneWayFeeText);
	
						    if (getOneWayFeeText.contains("ONE WAY FEE OF"))
						    {
								driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
								Thread.sleep(3000);
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
								Thread.sleep(10000);
							}
						}
						
						if (Counters.isDisplayed())
						{
						 	driver.findElement(By.id("checkinRepromptDialog:checkInProductRepromptForm:cpdYesBtn")).click();
							Thread.sleep(10000);
						}
						
						WebElement OneWayFee1 = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
						if (OneWayFee1.isDisplayed())
						{
							System.out.println("Property Identified inside Oneway Loop");
						    String getOneWayFeeText = OneWayFee1.getText();
						    System.out.println(" One way fee get text value " + getOneWayFeeText);
	
						    if (getOneWayFeeText.contains("ONE WAY FEE OF"))
						    {
								driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
								Thread.sleep(3000);
						    }
						}
						
					//Enter the values in SB screen
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:paymentMethod']")).sendKeys(MOP);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:ccType']")).sendKeys(CARDTYPE);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:ccNumber']")).sendKeys(CARDNO);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:ccExpireMonth']")).sendKeys(EXPMONTH);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:ccExpireYear']")).sendKeys(EXPYEAR);
						
						if ((driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:paymentReason")).isDisplayed() && driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:paymentReason")).isEnabled()))
						{
							driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSwipe:paymentReason")).sendKeys("No reader/reader not working");
						}
						
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:checkoutAuth']")).sendKeys(CKECKINOUT_AUTH);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:checkinAuth']")).sendKeys(CKECKINOUT_AUTH);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:address1']")).sendKeys(ADDR1);
						driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:address3']")).sendKeys(ADDR3);
						
						WebElement TotBill 	= driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingTotalAmount']"));
						System.out.println(TotBill.getText());
						String FirstDollar = TotBill.getText();
						if (FirstDollar.equals("$"))
						{	
							System.out.println(FirstDollar);
							test = extent.createTest(TESTCASENAME);
							System.out.println(TESTCASENAME);
	
								rwe.setCellData("CheckIn_Avis", K, 21, "Dollar Amount and number of Days values are not displayed");
								rwe.setCellData("CheckIn_Avis", K, 22, "FAIL");
								test.log(Status.FAIL, "Fail");
	
								Date d2= new Date();
								SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
								File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
								FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_Values-Not-Displayed-In-SplitBilling"+".pgn"));
	
							driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingCancelButton")).click();
							
							functions.logout();
							driver.close();
							
							
						}
						
						WebElement TotLor 	= driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingTotalDays"));
						System.out.println(TotLor.getText());
						String TotLorVal 	= TotLor.getText();
						rwe.setCellData("CheckIn_Avis", K, 19, TotLorVal);
						int ConvInt 		= Integer.parseInt(TotLorVal);
						System.out.println(ConvInt);
						
						System.out.println(TotBill.getText());
						String TotBillVal	= TotBill.getText();
						TotBillVal = TotBillVal.substring(1, TotBillVal.length()-3);
						System.out.println(TotBillVal);
						rwe.setCellData("CheckIn_Avis", K, 20, TotBillVal);
						
						int ContDLA 		= Integer.parseInt(TotBillVal);
						
						if (SBTYPE.equals("NOD"))
						{
														 
							driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingNumberOfDays']")).click();
							if(ConvInt > 1) {
									int ConvInt1;
									ConvInt1 = ConvInt/2;
									String FinlLor = "" +ConvInt1;
									driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSecondCardDays']")).sendKeys(FinlLor);
									}
							else{
									String FinlLor2 = "" +ConvInt;
									driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSecondCardDays']")).sendKeys(FinlLor2);
						
							}
						}
						else{
							
							int TotAmt;
							TotAmt	  = ContDLA/2;
							String FinlBill = "" +TotAmt;
							System.out.println(FinlBill);
							driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingDollarAmount']")).click();
							driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSecondCardAmount']")).sendKeys(FinlBill);
						}
							driver.findElement(By.xpath("//*[@id='menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingSaveButton']")).click();
							Thread.sleep(15000);
							
						WebElement Error 	= driver.findElement(By.xpath("//*[@id='checkinErrorDialog:checkInErrorForm:errorOut']"));
							
						if (Error.isDisplayed())
						{
							System.out.println("Error Code: -1");
						    String getErrorText = Error.getText();
						    System.out.println(" Error Code: -1 " + getErrorText);
	
						    if (getErrorText.contains("Error Code"))
						    {
								rwe.setCellData("CheckIn_Avis", K, 21, "Error Code: -1 displayed");
								rwe.setCellData("CheckIn_Avis", K, 22, "FAIL");
								test.log(Status.FAIL, "Fail");
	
								Date d2= new Date();
								SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
								File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
								FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_Error-Code"+".pgn"));
	
						    	driver.findElement(By.id("checkinErrorDialog:checkInErrorForm:errorOkButton")).click();
								driver.findElement(By.id("menulist:checkinContainer:checkInSplitBilling:splitBillingForm:splitBillingCancelButton")).click();
								functions.logout();
								driver.close();
						    }
						}
							
							WebElement strInvalidCard = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
							if (strInvalidCard.isDisplayed())
							{
							    String getstrInvalidCardText = strInvalidCard.getText();
							    System.out.println(" Invalid Card Info " + getstrInvalidCardText);
							    if (getstrInvalidCardText.contains("INVALID CREDIT CARD"))
							    {
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:4:repromptValue")).sendKeys("OK/12345");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(20000);
							    }
							}	
											
		
				 //-----(1) Is RSN code message box displayed
							WebElement Swipe = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
							if (Swipe.isDisplayed())
							{
							    String getReasonCode = Swipe.getText();
							    System.out.println("CARD SWIPE OR ENTER RSN CODE" + getReasonCode);
							    if (getReasonCode.contains("PLEASE OBTAIN CARD SWIPE OR ENTER RSN CODE"))
							    {
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(20000);
		
				//-----(2) Is RSN code message box displayed					    					    
							if (Swipe.isDisplayed())
							{
							    System.out.println("CARD SWIPE OR ENTER RSN CODE" + getReasonCode);
							    if (getReasonCode.contains("PLEASE OBTAIN CARD SWIPE OR ENTER RSN CODE"))
							    {
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(20000);
								
				//----When Invalid card message box displayed			
							if (driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut")).isDisplayed());
							{
								WebElement InvalidCard = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
								String getInvalidCard = InvalidCard.getText();
							    System.out.println("INVALID CARD" + getInvalidCard);
							    if (getInvalidCard.contains("INVALID CREDIT CARD"))
							    {
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:4:repromptValue")).sendKeys("OK/12345");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
										    	
											    }
											}
									    }
									}
							}  } 
	
				 //-----(1) Is RSN code message box displayed
		//					WebElement Swipe = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
							if (Swipe.isDisplayed())
							{
							    String getReasonCode = Swipe.getText();
							    System.out.println("CARD SWIPE OR ENTER RSN CODE" + getReasonCode);
							    if (getReasonCode.contains("PLEASE OBTAIN CARD SWIPE OR ENTER RSN CODE"))
							    {
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
							    	driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
							    	Thread.sleep(20000);
							    }
							}
							
					//-----(3) Is RSN code message box displayed					    					    
		//					WebElement strInvalidCard = driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptOut"));
							if (strInvalidCard.isDisplayed())
							{
							    String getstrInvalidCardText = strInvalidCard.getText();
							    System.out.println(" Invalid Card Info " + getstrInvalidCardText);
							    if (getstrInvalidCardText.contains("INVALID CREDIT CARD"))
							    {
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:2:repromptValue")).sendKeys("R");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptTable:4:repromptValue")).sendKeys("OK/12345");
									driver.findElement(By.id("checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton")).click();
									Thread.sleep(20000);
							    }
							}	
							
											
							WebElement Print = driver.findElement(By.xpath("//*[@id='printPreviewDialog:printPreviewForm:closePrintPreviewBtn']/span"));
							if (Print.isDisplayed())
							{
							    String getPrintText = Print.getText();
							    System.out.println(" Print Preview " + getPrintText);
			    
									Date d1= new Date();
									SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
									File src1= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
									FileUtils.copyFile(src1,new File( ScreenShotPath +TESTCASENAME+ sdf1.format(d1)+"_Print"+".pgn"));
								
							    	driver.findElement(By.xpath("//*[@id='printPreviewDialog:printPreviewForm:closePrintPreviewBtn']/span")).click();
							    	Thread.sleep(3000);
							}
							
							
					//Complete Check In
							
							WebElement CompCheckIn = driver.findElement(By.xpath("//*[@id='checkInSuccessDlg:checkInSuccessForm:checkInSuccessPanel_content']/div[1]/h5"));
							if (CompCheckIn.isDisplayed());
							{
							    String getCheckInText = CompCheckIn.getText();
							    System.out.println(" Complated CheckIn " + getCheckInText);
							    if (getCheckInText.contains("Complete Check In"))
							    {
								
							    	String CCI = CompCheckIn.getText();
									rwe.setCellData("CheckIn_Avis", K, 21, CCI);
									
									Date d2= new Date();
									SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
									File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
									FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_CheckInCompleted"+".pgn"));
		
									
							    	driver.findElement(By.xpath("//*[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).click();
							    	Thread.sleep(3000);
							    }
							
					
					// Logout the application
							test = extent.createTest(TESTCASENAME);
							
							if (rwe.getCellData("CheckIn_Avis", K, 21).isEmpty())
							{
								rwe.setCellData("CheckIn_Avis", K, 22, "FAIL");
								test.log(Status.FAIL, "Fail");
							}
							else
							{
								test.log(Status.PASS, "Pass");
								rwe.setCellData("CheckIn_Avis", K, 22, "PASS");
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