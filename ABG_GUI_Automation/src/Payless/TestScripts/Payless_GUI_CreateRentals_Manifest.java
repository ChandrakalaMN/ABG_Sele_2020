package Payless.TestScripts;

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
* '## SCRIPT NAME: GUI_CreateRentals_Manifest_Avis '## BRAND: AVIS '## DESCRIPTION:
* Cancel Reservation for already created Reservations
* products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
* All the required Test Data should be available in Test Data Sheet. '##
* OUTPUT: Rentals should be created successfully from the Manifest Screen
* NOTE: In the Manifest Screen data for teh current date + two days can be pulled 
* Date Created: 26 Sep 2018
* HISTORY 05-SEP-2018 - GUIFunctions class created for GUI Common
* functionalities and CR functionality
* '#############################################################################################################################
**/

public class Payless_GUI_CreateRentals_Manifest
{
	private static final String NULL = null;

	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport()
	{
		extent = Extentmanager.GetExtent();
	}

	// public static void main(String[] args) throws IOException, Exception,
	// FileNotFoundException

	@Test
	public void test() throws Exception {

		try {
			// Read input from excel
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Selenium\\Test_Data\\PAYLESS\\Payless_GUITestData_Rental_CheckOut_Manifest.xlsx");
			int intRowCount = rwe.intRowCount;

			System.out.println(" test data row count is " + intRowCount);

			for (int k = 0; k <= intRowCount; k++) {

				String Execute = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 2);
				if (Execute.equals("Y")) {
					System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\ChromeDriver\\chromedriver.exe");
					WebDriver driver = new ChromeDriver();
					GUIFunctions functions = new GUIFunctions(driver);

					driver.manage().window().maximize();
					driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
					System.out.println(" iteration " + k);
					String ptestcasename = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 4);
					String pclientURL = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 7);
					String poutSTA = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 10);
					String pthinClient = pclientURL + poutSTA;
					String puName = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 8);
					String ppswd = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 9);
					String pResNo = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 11);
					String pFromDate = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 12);
					String pFromTime = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 13);
					String pToDate = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 14);
					String pToTime = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 15);
					String pCustType = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 16);
					String pWizardNo = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 17);
					String pLastName = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 18);
					String pFirstName = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 19);
					String pCOCountry = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 20);
					String pCOState = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 21);
					String pCODRLICNO = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 22);
					String pCODOB = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 23);
					String pCOCOMPANY = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 24);
					String pCOADDR1 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 25);
					String pCOADDR2 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 26);
					String pCOADDR3 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 27);
					String pCOCONTACTINFO = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 28);
					String pCOEMAIL = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 29);
					String pCOMOP = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 30);
					String pCOCCDC = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 31);
					String pCARDNO = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 32);
					String pCOMM = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 33);
					String pCOYY = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 34);
					String pMOPREASON = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 35);
					String pRETURNSTATION = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 36);
					String pRETURNDATE = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 37);
					String pRETURNTIME = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 38);
					String pCARGROUP = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 39);
					String pAWD = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 40);
					String pCOUPON = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 41);
					String pFTNTYPE = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 42);
					String pFTNUMBER = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 43);
					String pREMARKS = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 44);
					String pCOMVA = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 45);
					String pCOMILEAGE = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 46);
					String pINSURANCE = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 47);
					String pCOUNTERPRODUCT = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 48);
					String pADRLASTNAME = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 49);
					String pADRFIRSTNAME = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 50);
					String pADRDOB = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 51);
					String pADRCOUNTRY = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 52);
					String pADRSTATE = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 53);
					String pADRDRLICNO = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 54);
					String pADREXPMM = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 55);
					String pADREXPYY = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 56);
					String pADRADDR1 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 57);
					String pADRADDR2 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 58);
					String pADRADDR3 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 59);
					String pADRTELEPHONE1 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 60);
					String pADRTELEPHONE2 = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 61);
					String pADRCCDC = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 62);
					String pADRCARDNO = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 63);
					String pADRCCEXPMM = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 64);
					String pADRCCEXPYY = rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 65);

					driver.get(pthinClient);
					functions.login(puName, ppswd);
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
					driver.findElement(By.xpath("//input[@id='manifestFromDate']")).sendKeys(pFromDate);
					driver.findElement(By.xpath("//input[@id='fromTime']")).click();
					driver.findElement(By.xpath("//input[@id='fromTime']")).clear();
					driver.findElement(By.xpath("//input[@id='fromTime']")).sendKeys(pFromTime);
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).click();
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).clear();
					driver.findElement(By.xpath("//input[@id='manifestToDate']")).sendKeys(pToDate);
					driver.findElement(By.xpath("//input[@id='toTime']")).click();
					driver.findElement(By.xpath("//input[@id='toTime']")).clear();
					driver.findElement(By.xpath("//input[@id='toTime']")).sendKeys(pToTime);
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
								System.out.println("The selected Reservation number is " + pResNo);
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
								if (pResNo.equals(Reservation))
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
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys(pCOCountry);
										System.out.println("Entered the Country");
									}

									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).click();
									String State = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).getAttribute("value");
									System.out.println("The state is " + State);
									if (State.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys(pCOState);
										System.out.println("Entered the State");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).click();
									String Licence = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).getAttribute("value");
									System.out.println("The Licence entered id " + Licence);
									if (Licence.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys(pCODRLICNO);
										System.out.println("Entered the Licence Number");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).click();
									String DOB = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).getAttribute("value");
									System.out.println("The Date od Birth entered is " + DOB);
									if (DOB.isEmpty()) 
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys(pCODOB);
										System.out.println("Entered the Date of Birth");
									}

									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).click();
									String Company = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).getAttribute("value");
									System.out.println("The Company entered is " + Company);
									if (Company.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys(pCOCOMPANY);
										System.out.println("Entered the Company");
									}
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).click();
									String Add1 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).getAttribute("value");
									if (Add1.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys(pCOADDR1);
										System.out.println("Entered the Address 1");
									}
									String Add2 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).getAttribute("value");
									if (Add2.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys(pCOADDR2);
										System.out.println("Entered the Address 2");
									}
									String Add3 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).getAttribute("value");
									if (Add3.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys(pCOADDR3);
										System.out.println("Entered the Address 3");
									}
									String Contact = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).getAttribute("value");
									if (Contact.isEmpty()) 
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys(pCOCONTACTINFO);
										System.out.println("Entered the Contact Info");
									}
									if (driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).isEnabled()) 
									{
										String Email = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).getAttribute("value");
										if (Email.isEmpty()) 
										{
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys(pCOEMAIL);
											System.out.println("Entered the Email");
										}
									}
									if (pFTNTYPE != NULL) 
									{
										String ftntp = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).getAttribute("value");
										if (ftntp.isEmpty())
										{
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).sendKeys(pFTNTYPE);
											System.out.println("Entered the FTN Type");
											driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftNumber']")).sendKeys(pFTNUMBER);
											System.out.println("Entered the FTN Number");
										}
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:verifiedFTN']")).click();
										System.out.println("Verified the FTN number");
									}

									// driver.findElement(By.xpath("//input[@id='glyphicon
									// glyphicon-plus']")).click();

									Select s = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:payMethod")));
									s.selectByValue(pCOMOP);
									System.out.println("Entered the Method of Payment");
									Select s2 = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType")));
									s2.selectByValue(pCOCCDC);
									System.out.println("Entered the Type of Payment");
									String CardNumber = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).getAttribute("value");
									if (CardNumber.isEmpty()) 
									{

										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).sendKeys(pCARDNO);
										System.out.println("Entered the Card Number");
									}
									String Month = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).getAttribute("value");
									if (Month.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).sendKeys(pCOMM);
										System.out.println("Entered the Expiry Month");
									}
									String year = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).getAttribute("value");
									if (year.isEmpty())
									{
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).sendKeys(pCOYY);
										System.out.println("Entered the Expiry Year");
									}

									Select s3 = new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason")));
									s3.selectByValue(pMOPREASON);
									System.out.println("Entered the Reason for payment");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:mopCCI']")).sendKeys("1234");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:authorization']")).sendKeys("OK/1234");
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).clear();
									// driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).sendKeys("GB/C");
									String RateCode = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).getAttribute("value");
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 66, RateCode);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(pCOMVA);
									System.out.println("Entered the MVA number");
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
									Thread.sleep(2000);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
									Thread.sleep(2000);
									driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(pCOMILEAGE);
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
											rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 74, OptMSG);
											rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 75, "FAIL");
											test.log(Status.FAIL, "Fail");
											String ScreenShotPath = "C:\\Selenium\\Screenshots\\Payless\\Payless_Create_Rentals_Manifest\\";
											functions.ScreenCapture(ScreenShotPath, ptestcasename);
											System.out.println("The Screenshot is taken");
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton"))
													.click();
											driver.close();
											break;
										} 
										else if (OptMSG.equals("START DATE INVALID FOR RATE")) 
										{
											System.out.println("ERROR MESSAGE " + OptMSG + " is Displayed");
											rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 74, OptMSG);
											rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 75, "FAIL");
											test.log(Status.FAIL, "Fail");
											String ScreenShotPath = "C:\\Selenium\\Screenshots\\Payless\\Payless_Create_Rentals_Manifest\\";
											functions.ScreenCapture(ScreenShotPath, ptestcasename);
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
												if (pCOCCDC.equals("CA")) 
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
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 67, strCOLNFNGetText);
									System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
									String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 68, strCOVehicleModelGetText);
									System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
									String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 69, strCOResNoGetText);
									System.out.println(" Reservation No value is " + strCOResNoGetText);
									String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 70, strCOMVAGetText);
									System.out.println(" MVA No value is " + strCOMVAGetText);
									String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 71, strCORANoGetText);
									System.out.println(" RA Number value is " + strCORANoGetText);
									String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 72, strCOLicensePlateNoGetText);
									System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
									String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 73, strCOEstimatedCostGetText);
									System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
									String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
									Thread.sleep(7000);
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 74, strCOSystemMsgGetText);
									System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
									Thread.sleep(7000);

									String ScreenShotPath = "C:\\Selenium\\Screenshots\\Payless\\Payless_Create_Rentals_Manifest\\";
									functions.ScreenCapture(ScreenShotPath, ptestcasename);
									System.out.println("The Screenshot is taken");
									Thread.sleep(5000);
									if (driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).isDisplayed()) 
									{
										driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
										System.out.println("Clicked on Done Button");
										Thread.sleep(7000);
									}
									String ScreenShotPath2 = "C:\\Selenium\\Screenshots\\Payless\\Payless_Create_Rentals_Manifest\\";
									functions.ScreenCapture(ScreenShotPath2, ptestcasename);
									System.out.println("The Screenshot is taken");
									driver.close();
									Thread.sleep(2000);
									// functions.closeWindows();
									test = extent.createTest(ptestcasename);
									if (rwe.getCellData("Payless_GUI_Rentals_Manifest", k, 68).isEmpty()) 
									{
										rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 75, "FAIL");
										test.log(Status.FAIL, "Fail");
									} else 
									{
										rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 75, "PASS");
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
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 74, "The Required Reservation number is not available in the Manifest Screen");
									rwe.setCellData("Payless_GUI_Rentals_Manifest", k, 75, "FAIL");
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
