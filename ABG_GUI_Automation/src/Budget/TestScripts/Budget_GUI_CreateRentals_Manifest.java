package Budget.TestScripts;

import java.util.concurrent.TimeUnit;

//import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;

import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;
/* '#############################################################################################################################
* '## SCRIPT NAME: Budget_GUI_CreateRentals_Manifest '## BRAND: BUDGET '## DESCRIPTION:
* Cancel Reservation for already created Reservations
* products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
* All the required Test Data should be available in Test Data Sheet. '##
* OUTPUT: Rentals should be created successfully from the Manifest Screen
* NOTE: In the Manifest Screen data for the current date + two days can be pulled 
* Date Created: 15 Oct 2018
* HISTORY 05-SEP-2018 - GUIFunctions class created for GUI Common
* functionalities and CR functionality
* '#############################################################################################################################
**/

public class Budget_GUI_CreateRentals_Manifest
{
	private static final String NULL = null;

	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport() {

		extent = Extentmanager.GetExtent();
	}

	// public static void main(String[] args) throws IOException, Exception,
	// FileNotFoundException

	@Test
	public void test() throws Exception {

		try {
			// Read input from excel
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Selenium\\Test_Data\\BUDGET\\Budget_GUITestData_Rental_CheckOut_Manifest.xlsx");
			int intRowCount = rwe.intRowCount;

			System.out.println(" test data row count is " + intRowCount);

			for (int k = 0; k <= intRowCount; k++) {

				String Execute = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 2);
				if (Execute.equals("Y")) {
					System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\ChromeDriver\\chromedriver.exe");
					WebDriver driver = new ChromeDriver();
					GUIFunctions functions = new GUIFunctions(driver);

					driver.manage().window().maximize();
					driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
					System.out.println(" iteration " + k);
					String btestcasename = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 4);
					String bclientURL = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 7);
					String boutSTA = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 10);
					String bthinClient = bclientURL + boutSTA;
					String buName = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 8);
					String bpswd = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 9);
					String bResNo = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 11);
					String bFromDate = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 12);
					String bFromTime = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 13);
					String bToDate = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 14);
					String bToTime = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 15);
					String bCustType = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 16);
					String bWizardNo = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 17);
					String bLastName = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 18);
					String bFirstName = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 19);
					String bCOCountry = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 20);
					String bCOState = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 21);
					String bCODRLICNO = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 22);
					String bCODOB = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 23);
					String bCOCOMPANY = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 24);
					String bCOADDR1 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 25);
					String bCOADDR2 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 26);
					String bCOADDR3 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 27);
					String bCOCONTACTINFO = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 28);
					String bCOEMAIL = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 29);
					String bCOMOP = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 30);
					String bCOCCDC = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 31);
					String bCARDNO = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 32);
					String bCOMM = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 33);
					String bCOYY = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 34);
					String bMOPREASON = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 35);
					String bRETURNSTATION = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 36);
					String bRETURNDATE = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 37);
					String bRETURNTIME = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 38);
					String bCARGROUP = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 39);
					String bAWD = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 40);
					String bCOUPON = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 41);
					String bFTNTYPE = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 42);
					String bFTNUMBER = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 43);
					String bREMARKS = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 44);
					String bCOMVA = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 45);
					String bCOMILEAGE = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 46);
					String bINSURANCE = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 47);
					String bCOUNTERPRODUCT = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 48);
					String bADRLASTNAME = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 49);
					String bADRFIRSTNAME = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 50);
					String bADRDOB = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 51);
					String bADRCOUNTRY = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 52);
					String bADRSTATE = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 53);
					String bADRDRLICNO = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 54);
					String bADREXPMM = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 55);
					String bADREXPYY = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 56);
					String bADRADDR1 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 57);
					String bADRADDR2 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 58);
					String bADRADDR3 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 59);
					String bADRTELEPHONE1 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 60);
					String bADRTELEPHONE2 = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 61);
					String bADRCCDC = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 62);
					String bADRCARDNO = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 63);
					String bADRCCEXPMM = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 64);
					String bADRCCEXPYY = rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 65);

					driver.get(bthinClient);
					functions.login(buName, bpswd);
					Thread.sleep(5000);

					String strPageTitle = driver.getTitle();

					if (strPageTitle.equals("Avis Budget Group")) {
						System.out.println(" QA.user logging in... ");

						System.getProperty("user.dir");
					} else {
						driver.navigate().refresh();
					}

					driver.findElement(By.xpath("//a[@data-target='#menulist\\:manifestlink']")).click();
					Thread.sleep(10000);
					driver.findElement(By.xpath("//input[@id='manifestFromDate']")).click();
					driver.findElement(By.xpath("//input[@id='manifestFromDate']")).clear();
					driver.findElement(By.xpath("//input[@id='manifestFromDate']")).sendKeys(bFromDate);
					driver.findElement(By.xpath("//input[@id='fromTime']")).click();
					driver.findElement(By.xpath("//input[@id='fromTime']")).clear();
					driver.findElement(By.xpath("//input[@id='fromTime']")).sendKeys(bFromTime);
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).click();
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).clear();
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).sendKeys(bToDate);
					driver.findElement(By.xpath("//input[@id='toTime']")).click();
					driver.findElement(By.xpath("//input[@id='toTime']")).clear();
					driver.findElement(By.xpath("//input[@id='toTime']")).sendKeys(bToTime);
					driver.findElement(By.xpath("//button[@id='getManifest']")).click();
					Thread.sleep(3000);
					
					   WebElement TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
				       String TotItems1= TotItems.getText();
				       String str = TotItems1;
				       
				       //Trim spaces and fetch only tot value
				       str = str.substring(9, str.length()-5);
				       System.out.println("The String value is "+str);
				       int a= 0;
				       int Row = Integer.parseInt(str.trim());
				       System.out.println("The value of Row is " + Row);
				       int RowValue = a + Row ;
				       
				       System.out.println("The RowValue is "+RowValue);
				       
				       Thread.sleep(2000);
					WebElement table1 = driver.findElement(By.xpath("//div[@id='manifestGrid']"));
					if (table1.isDisplayed()) {
						String Sess_ID = table1.getAttribute("className");
						System.out.println(Sess_ID);
						Sess_ID = Sess_ID.replace("manifest-grid ui-grid ng-isolate-scope grid", "");
						for (int i = 0; i < RowValue; i++)

						{
							String Dy_xpath = "//*[@id=\"" + Sess_ID + "-" + i + "-uiGrid-003A-cell\"]/div";
							System.out.println("Dy_xpath is:" + Dy_xpath);
							Thread.sleep(10000);
							try {
								table1.findElement(By.xpath(Dy_xpath)).click();
								System.out.println(table1.findElement(By.xpath(Dy_xpath)).getText());
								driver.findElement(By.xpath("//button[@ng-click='viewResDetail()']")).click();
								Thread.sleep(10000);
								String Reservation = driver.findElement(By.xpath("//span[@ng-bind='res.reservationNum']")).getText();
								System.out.println("The Reservation numbers is " + Reservation);
								System.out.println("The selected Reservation number is " + bResNo);
								Thread.sleep(10000);
								System.out.println("Clicked on the  " + i + "th row");
								/*
								 * if ( ResNo != Reservation) {
								 * driver.findElement(By.xpath(
								 * "//input[@id='mdetailCloseButton']")).click()
								 * ; System.out.
								 * println("The Reservation displayed does not match the entered Reservation number"
								 * ); Thread.sleep(5000); }
								 */
								String RowValue1;
								if (bResNo.equals(Reservation))
								{
									
									driver.findElement(By.xpath("//input[@id='resDetailCheckout']")).click();
									Thread.sleep(3000);
									System.out.println("The Reservation is displayed in the Check Out Screen");
									/*driver.findElement(By.xpath("//input[@id='searchCommandLink']")).click();
									System.out.println("Clicked on the Search Button");*/
									Thread.sleep(20000);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).click();
								
									String Country = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).getAttribute("value");
									System.out.println("The country is " + Country);
									if (Country.isEmpty()) 
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys(bCOCountry);
										System.out.println("Entered the Country");
									}

									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).click();
									String State = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).getAttribute("value");
									System.out.println("The state is " + State);
									if (State.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys(bCOState);
										System.out.println("Entered the State");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).click();
									String Licence = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).getAttribute("value");
									System.out.println("The Licence entered id " + Licence);
									if (Licence.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys(bCODRLICNO);
										System.out.println("Entered the Licence Number");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).click();
									String DOB = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).getAttribute("value");
									System.out.println("The Date od Birth entered is " + DOB);
									if (DOB.isEmpty()) 
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys(bCODOB);
										System.out.println("Entered the Date of Birth");
									}

									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).click();
									String Company = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).getAttribute("value");
									System.out.println("The Company entered is " + Company);
									if (Company.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys(bCOCOMPANY);
										System.out.println("Entered the Company");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).click();
									String Add1 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).getAttribute("value");
									if (Add1.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys(bCOADDR1);
										System.out.println("Entered the Address 1");
									}
									String Add2 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).getAttribute("value");
									if (Add2.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys(bCOADDR2);
										System.out.println("Entered the Address 2");
									}
									String Add3 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).getAttribute("value");
									if (Add3.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys(bCOADDR3);
										System.out.println("Entered the Address 3");
									}
									String Contact = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).getAttribute("value");
									if (Contact.isEmpty()) 
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys(bCOCONTACTINFO);
										System.out.println("Entered the Contact Info");
									}
									if (driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).isEnabled()) 
									{
										String Email = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).getAttribute("value");
										if (Email.isEmpty()) 
										{
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys(bCOEMAIL);
											System.out.println("Entered the Email");
										}
									}
									if (bFTNTYPE != NULL) 
									{
										String ftntp = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).getAttribute("value");
										if (ftntp.isEmpty())
										{
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).sendKeys(bFTNTYPE);
											System.out.println("Entered the FTN Type");
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftNumber']")).sendKeys(bFTNUMBER);
											System.out.println("Entered the FTN Number");
										}
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:verifiedFTN']")).click();
										System.out.println("Verified the FTN number");
									}

									// driver.findElement(By.xpath("//input[@id='glyphicon
									// glyphicon-plus']")).click();

									Select s = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:payMethod")));
									s.selectByValue(bCOMOP);
									System.out.println("Entered the Method of Payment");
									Select s2 = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType")));
									s2.selectByValue(bCOCCDC);
									System.out.println("Entered the Type of Payment");
									String CardNumber = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).getAttribute("value");
									if (CardNumber.isEmpty()) 
									{

										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).sendKeys(bCARDNO);
										System.out.println("Entered the Card Number");
									}
									String Month = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).getAttribute("value");
									if (Month.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).sendKeys(bCOMM);
										System.out.println("Entered the Expiry Month");
									}
									String year = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).getAttribute("value");
									if (year.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).sendKeys(bCOYY);
										System.out.println("Entered the Expiry Year");
									}

									Select s3 = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason")));
									s3.selectByValue(bMOPREASON);
									System.out.println("Entered the Reason for payment");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:mopCCI']")).sendKeys("1234");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:authorization']")).sendKeys("OK/1234");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).clear();
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).sendKeys("GB/C");
									String RateCode = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).getAttribute("value");
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 66, RateCode);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(bCOMVA);
									System.out.println("Entered the MVA number");
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
									Thread.sleep(2000);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
									Thread.sleep(2000);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(bCOMILEAGE);
									System.out.println("Entered the Mileage");
									if (driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).isDisplayed()) 
									{
										driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).click();
									}
									Thread.sleep(7000);
									if (driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).isDisplayed()) 
									{
										String OptMSG = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getAttribute("value");
										System.out.println("The Output message in vechicle continue screen is " + OptMSG);
										if (OptMSG.equals("NET RATE ERROR")) 
										{
											System.out.println("ERROR MESSAGE " + OptMSG + " is Displayed");
											rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 74, OptMSG);
											rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 75, "FAIL");
											test.log(Status.FAIL, "Fail");
											String ScreenShotPath = "C:\\Selenium\\Screenshots\\Budget\\Budget_Create_Rentals_Manifest\\";
											functions.ScreenCapture(ScreenShotPath, btestcasename);
											System.out.println("The Screenshot is taken");
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton"))
													.click();
											driver.close();
											break;
										} 
										else if (OptMSG.equals("START DATE INVALID FOR RATE")) 
										{
											System.out.println("ERROR MESSAGE " + OptMSG + " is Displayed");
											rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 74, OptMSG);
											rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 75, "FAIL");
											test.log(Status.FAIL, "Fail");
											String ScreenShotPath = "C:\\Selenium\\Screenshots\\Budget\\Budget_Create_Rentals_Manifest\\";
											functions.ScreenCapture(ScreenShotPath, btestcasename);
											System.out.println("The Screenshot is taken");
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton"))
													.click();
											driver.close();
											break;

										}
									}

									Thread.sleep(20000);
									if (driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).isDisplayed())
									{
										driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
										System.out.println("Clicked on Continue Button");
										Thread.sleep(20000);
									}
									Thread.sleep(20000);
									if (driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).isDisplayed())
									{
										driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).click();
										System.out.println("Clicked on Complete Checkout Button");
										Thread.sleep(7000);
									}

									Thread.sleep(7000);
									try 
									{

										if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())

										{
											String str1 = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
											System.out.println("The message displayed is " + str1);
											if (str1.contains("Please enter the Credit ID security code (CVV/CCV)")) 
											{
												if (bCOCCDC.equals("CA")) 
												{
													driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("1234");
													driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
													Thread.sleep(7000);

												} 
												else
												{
													driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("123");
													driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
													Thread.sleep(7000);

												}
												if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed()) 
												{
													String str2 = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
													System.out.println("The message displayed is " + str2);
													if (str2.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION")) 
													{
														driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
														driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
														driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
														Thread.sleep(4000);
													}
												}

											} 
											else 
											if (str1.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION ")) 
											{
												driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
												driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
												driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
												Thread.sleep(4000);
											}
										}
									} catch (Exception e)
									{
										e.printStackTrace();
									}

									// Checkout complete Screen and update in Excel

									Thread.sleep(7000);
									String strCOLNFNGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 67, strCOLNFNGetText);
									System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
									String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 68, strCOVehicleModelGetText);
									System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
									String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 69, strCOResNoGetText);
									System.out.println(" Reservation No value is " + strCOResNoGetText);
									String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 70, strCOMVAGetText);
									System.out.println(" MVA No value is " + strCOMVAGetText);
									String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 71, strCORANoGetText);
									System.out.println(" RA Number value is " + strCORANoGetText);
									String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 72, strCOLicensePlateNoGetText);
									System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
									String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 73, strCOEstimatedCostGetText);
									System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
									String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 74, strCOSystemMsgGetText);
									System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
									Thread.sleep(7000);

									String ScreenShotPath = "C:\\Selenium\\Screenshots\\Budget\\Budget_Create_Rentals_Manifest\\";
									functions.ScreenCapture(ScreenShotPath, btestcasename);
									System.out.println("The Screenshot is taken");
									Thread.sleep(5000);
									if (driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).isDisplayed()) 
									{
										driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
										System.out.println("Clicked on Done Button");
										Thread.sleep(7000);
									}
									String ScreenShotPath2 = "C:\\Selenium\\Screenshots\\Budget\\Budget_Create_Rentals_Manifest\\";
									functions.ScreenCapture(ScreenShotPath2, btestcasename);
									System.out.println("The Screenshot is taken");
									driver.close();
									Thread.sleep(2000);
									// functions.closeWindows();
									test = extent.createTest(btestcasename);
									if (rwe.getCellData("Budget_GUI_Rentals_Manifest", k, 68).isEmpty()) 
									{
										rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 75, "FAIL");
										test.log(Status.FAIL, "Fail");
									} else 
									{
										rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 75, "PASS");
										test.log(Status.PASS, "Pass");
									}

									break;
								}
								
								else 
									RowValue1 = Integer.toString(RowValue-1);
									//System.out.println("The RowValue1 is "+RowValue1);
									String j = Integer.toString(i);
									//System.out.println("The value of j is"+ j);
									if(RowValue1.equals(j))
								   {
									
									System.out.println("The Reservation number required is not displayed in the Manifest Screen");
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 74, "The Required Reservation number is not available in the Manifest Screen");
									rwe.setCellData("Budget_GUI_Rentals_Manifest", k, 75, "FAIL");
									//test.log(Status.FAIL, "Fail");
									driver.findElement(By.xpath("//input[@id='mdetailCloseButton']")).click();
									driver.close();
									break;
									}

								else 
								    {
									
									driver.findElement(By.xpath("//input[@id='mdetailCloseButton']")).click();
									System.out.println("The Reservation displayed does not match the entered Reservation number");
									Thread.sleep(5000);
									continue;
									
								    }
								
							} catch (Exception e1) {
								e1.printStackTrace();
							}
							
						}

					}
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
