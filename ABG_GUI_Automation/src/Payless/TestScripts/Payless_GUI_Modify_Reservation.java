package Payless.TestScripts;

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

import AVIS.CommonFunctions.*;
import AVIS.CommonFunctions.ReadWriteExcel;

/**
* '#############################################################################################################################
* '## SCRIPT NAME: GUI_Modify_Reservation_Payless 
* '## BRAND: Payless 
* '## DESCRIPTION: Modify Reservation for all the editable fields 
* '## FUNCTIONAL AREA : Reservation Rates Screen 
* '## PRECONDITION:All the required Test Data should be available in Test Data Sheet. 
* '##OUTPUT: Reservation should be modified successfully.
* 
*HISTORY 12-SEP-2018 - GUIFunctions class created for GUI Common
* functionalities and CR functionality
* '#############################################################################################################################
**/

public class Payless_GUI_Modify_Reservation
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
	
	@SuppressWarnings("unlikely-arg-type")
	//public static void main(String[] args) throws IOException, Exception, FileNotFoundException {
	@Test
	public void test() throws Exception {
		try {
		// Read input from excel
		
		for (int k = 1; k <= 20; k++)
		{
			Payless_GUI_Modify_Reservation payless = new Payless_GUI_Modify_Reservation();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Users\\cmn\\Desktop\\Selenium\\TestData\\Payless_GUITestData_ModifyReservation.xlsx");
			String Execute = rwe.getCellData("Payless_GUI", k, 2);
			
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
				String testcasename    = rwe.getCellData("Payless_GUI", k, 4);
				String tokenURL    = rwe.getCellData("Payless_GUI", k, 6);
				String clientURL   = rwe.getCellData("Payless_GUI", k, 7);
				String outSTA      = rwe.getCellData("Payless_GUI", k, 8);
				String thinClient  = clientURL+outSTA;
				String uName       = rwe.getCellData("Payless_GUI", k, 9);
				String pswd        = rwe.getCellData("Payless_GUI", k, 10);
				String lstname     = rwe.getCellData("Payless_GUI", k, 11);
				String fstname     = rwe.getCellData("Payless_GUI", k, 12);
				String comonth		= rwe.getCellData("Payless_GUI", k, 13);	
				String codte       = rwe.getCellData("Payless_GUI", k, 14);
				String cotme       = rwe.getCellData("Payless_GUI", k, 15);
				String insta       = rwe.getCellData("Payless_GUI", k, 16);
				String cimonth =	rwe.getCellData("Payless_GUI", k, 17);
				String cidte       = rwe.getCellData("Payless_GUI", k, 18);
				String citme       = rwe.getCellData("Payless_GUI", k, 19);
				String carGrp      = rwe.getCellData("Payless_GUI", k, 20);
				String awd         = rwe.getCellData("Payless_GUI", k, 21);
				String cardname    = rwe.getCellData("Payless_GUI", k, 22);
				String cardNumber  = rwe.getCellData("Payless_GUI", k, 23);
				String expireMonth = rwe.getCellData("Payless_GUI", k, 24);
				String expireYear  = rwe.getCellData("Payless_GUI", k, 25);
				String reason  = rwe.getCellData("Payless_GUI", k, 26);
				String Ftncode      = rwe.getCellData("Payless_GUI", k, 27);
				String Ftnno      = rwe.getCellData("Payless_GUI", k, 28);
				String Reservation_No= rwe.getCellData("Payless_GUI", k, 29);
				String Modout = rwe.getCellData("Payless_GUI", k, 30);
				String Insurance = rwe.getCellData("Payless_GUI", k, 31);
				String Counterproducts = rwe.getCellData("Payless_GUI", k, 32);
				/* Open GUI URL's */
				driver.get(thinClient);
				
				
				
				
				/* Login */
				functions.login(uName, pswd);
				driver.navigate().refresh();
				functions.navigateToTab("ReservationRates");
				Thread.sleep(5000L);
				
				//Entering the reservation number in search field
				driver.findElement(By.xpath("//input[@placeholder='Display RA, Res, Wizard #, MVA #']")).click();
				driver.findElement(By.xpath("//input[@placeholder='Display RA, Res, Wizard #, MVA #']")).sendKeys(Reservation_No);
				driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:searchResForm:searchCommandLink']")).click();
				Thread.sleep(5000);
				
				//Modifying checkout location
				//String strNullValue = null;
				if(!Modout.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:pickupStation:pickupStation']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:pickupStation:pickupStation']")).clear();
					Thread.sleep(3000);
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:pickupStation:pickupStation']")).sendKeys(Modout);
					
				}
				
				//modifying checkin location
				else if(!insta.isEmpty()) 
				{
					
				driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:returnStation:returnStation']")).click();
				driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:returnStation:returnStation']")).clear();
				driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:returnStation:returnStation']")).sendKeys(insta);
				driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:returnStation:returnStation']")).sendKeys(Keys.TAB);
				Thread.sleep(3000);
				//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
				//Thread.sleep(3000);
				
				
				}
				//modifying firstname
				else if(!fstname.isEmpty())
				{
					driver.findElement(By.cssSelector("input[id='menulist:rateshopContainer:resForm:firstName']")).click();
					driver.findElement(By.cssSelector("input[id='menulist:rateshopContainer:resForm:firstName']")).clear();
					driver.findElement(By.cssSelector("input[id='menulist:rateshopContainer:resForm:firstName']")).sendKeys(fstname);
					Thread.sleep(3000);
					//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
					//Thread.sleep(3000);
					
				}
				//modifying lastname
				else if(!lstname.isEmpty())
				{
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:lastName']")).click();
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:lastName']")).clear();
					driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:lastName']")).sendKeys(lstname);
					Thread.sleep(3000);
					//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
					//Thread.sleep(3000);
					//driver.close();
				}	
					//modifying checkout date and time
					else if(!cimonth.isEmpty())
					{
						driver.findElement(By.xpath("//input[@id= 'menulist:rateshopContainer:resForm:checkout1_hid']")).click();
						driver.findElement(By.xpath("//input[@id= 'menulist:rateshopContainer:resForm:checkout1_hid']")).clear();
						
						while(!driver.findElement(By.cssSelector("[class='ui-datepicker-title'] [class='ui-datepicker-month']")).getText().contains(comonth))
						{
							driver.findElement(By.cssSelector("[class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all'] [class='ui-icon ui-icon-circle-triangle-e']")).click();
						}
						
						List<WebElement> dates = driver.findElements(By.xpath("//*[@data-handler='selectDay']"));
						
						int Count = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).size();
						
						for(int i=0;i<Count;i++)
						{
							String text = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).getText();
							if(text.equalsIgnoreCase(codte))
							{
								driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).click();
								break;
							}
						}
						//driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:lastName']")).sendKeys(codte);
						Thread.sleep(3000);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkout2']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkout2']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkout2']")).sendKeys(cotme);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkout2']")).sendKeys(Keys.ENTER);
						Thread.sleep(3000);
						//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
						//Thread.sleep(3000);
						
						//Modifying checkin date and time
						//else if((!cidte.isEmpty()) & (!citme.isEmpty()))
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin1_hid']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin1_hid']")).clear();
						//Picking checkin date and time
						while(!driver.findElement(By.cssSelector("[class='ui-datepicker-title'] [class='ui-datepicker-month']")).getText().contains(cimonth))
						{
							driver.findElement(By.cssSelector("[class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all'] [class='ui-icon ui-icon-circle-triangle-e']")).click();
						}
						
						List<WebElement> datescheckin = driver.findElements(By.xpath("//*[@data-handler='selectDay']"));
						
						int Count1 = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).size();
						
						for(int i=0;i<Count1;i++)
						{
							String text = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).getText();
							if(text.equalsIgnoreCase(cidte))
							{
								driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).click();
								break;
							}
						}
						//driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin1_hid']")).sendKeys(cidte);
						Thread.sleep(3000);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).sendKeys(citme);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).sendKeys(Keys.TAB);
						Thread.sleep(3000);
						//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
						//Thread.sleep(3000);
					}
				
					
					
						
					
					//modifying car group
					else if(!carGrp.isEmpty())
					{
						Select cargroup= new Select(driver.findElement(By.id("menulist:rateshopContainer:resForm:carGroup")));
						cargroup.selectByValue(carGrp);
						//Selecting the Ratecode on  Pop up window
						driver.findElement(By.xpath("//input[@id='searchCommandLinkResRateCode']")).click();
						//driver.findElement(By.xpath("//input[@name='radioRate']")).click();
						driver.findElement(By.xpath("//input[@value='Select Rate']")).click();
						//driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:rateCode']")).click();
						Thread.sleep(5000L);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:discountNumber']")).click();
						Thread.sleep(3000);
						//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
						//Thread.sleep(3000);
						
					}
				
				//Modifying the FTN number
					else if(!Ftnno.isEmpty())
					{
						driver.findElement(By.xpath("//span[@id='custToggle']")).click();
						Thread.sleep(3000);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:rftnType']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:rftnType']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:rftnType']")).sendKeys(Ftncode);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:ftNumber']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:ftNumber']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:ftNumber']")).sendKeys(Ftnno);
						Thread.sleep(3000);
					}
				//Modify Mop
				/* Enter MOP details */
					else if(!cardname.isEmpty()) {
						functions.expandPaymentInfoSection();
						functions.enterPaymentInformations(cardname, cardNumber, expireMonth, expireYear,reason);
					}
				//Modifying the counterproduct
				//functions.expandProtectionCoverageSection();
					if(Counterproducts.isEmpty()) {
					//if(Counterproducts.isEmpty()) {
						System.out.print("No CounterProduct selected");
						// break;
					} else {
						driver.findElement(By.xpath("//input[@value='Modify INS/CP']")).click();
						String[] cpVal = rwe.getCellData("Payless_GUI", k, 32).split("-");
						for (String e : cpVal) {
							WebDriverWait wait = new WebDriverWait(driver, 10);
							try {
								if (e.equalsIgnoreCase("ADR")) {
									WebElement cp1 = driver.findElement(By.id("productQuantity40"));
									Select cpADR = new Select(cp1);
									wait.until(ExpectedConditions.visibilityOf(cp1));
									if (cp1.isDisplayed()) {
										cpADR.selectByVisibleText("1");
									} else {
										break;
									}

								} else if (e.equalsIgnoreCase("CBS")) {
									WebElement cp2 = driver.findElement(By.id("productQuantity32"));
									Select cpCBS = new Select(cp2);
									wait.until(ExpectedConditions.visibilityOf(cp2));
									if (cp2.isDisplayed()) {
										cpCBS.selectByVisibleText("1");
									} else {
										break;
									}
								} else if (e.equalsIgnoreCase("CSS")) {
									WebElement cp3 = driver.findElement(By.id("productQuantity34"));
									Select cpCSS = new Select(cp3);
									wait.until(ExpectedConditions.visibilityOf(cp3));
									if (cp3.isDisplayed()) {
										cpCSS.selectByVisibleText("1");
									} else {
										break;
									}
								} else if (e.equalsIgnoreCase("GPS")) {
									WebElement cp4 = driver.findElement(By.id("productQuantityYesNo6"));
									Select cpGPS = new Select(cp4);
									wait.until(ExpectedConditions.visibilityOf(cp4));
									if (cp4.isDisplayed()) {
										cpGPS.selectByVisibleText("Y");
									} else {
										break;
									}
								} else if (e.equalsIgnoreCase("RSN")) {
									WebElement cp5 = driver.findElement(By.id("productQuantityYesNo10"));
									Select cpRSN = new Select(cp5);
									wait.until(ExpectedConditions.visibilityOf(cp5));
									if (cp5.isDisplayed()) {
										cpRSN.selectByVisibleText("Y");
									} else {
										break;
									}
								} else if (e.equalsIgnoreCase("TAB")) {
									WebElement cp6 = driver.findElement(By.id("productQuantityYesNo12"));
									Select cpTAB = new Select(cp6);
									wait.until(ExpectedConditions.visibilityOf(cp6));
									if (cp6.isDisplayed()) {
										cpTAB.selectByVisibleText("Y");
									} else {
										break;
									}
								} else if (e.equalsIgnoreCase("ESP")) {
									WebElement cp7 = driver.findElement(By.id("productQuantityYesNo6"));
									Select cpESP = new Select(cp7);
									wait.until(ExpectedConditions.visibilityOf(cp7));
									if (cp7.isDisplayed()) {
										cpESP.selectByVisibleText("Y");
									} else {
										break;
									}

								} else if (e.equalsIgnoreCase("SNB")) {
									WebElement cp8 = driver.findElement(By.id("productQuantityYesNo11"));
									Select cpSNB = new Select(cp8);
									if (cp8.isDisplayed()) {
										wait.until(ExpectedConditions.visibilityOf(cp8));
										cpSNB.selectByVisibleText("Y");
									} else {
										break;
									}
								}
							} catch (Exception e1) {
								e1.printStackTrace();
							}
						}

					}
				
				//Modifying the Insurance
					/* Add Insurances */
					Thread.sleep(5000);
					//driver.findElement(By.xpath("//input[@value='Modify INS/CP']")).click();
					functions.expandProtectionCoverageSection();
					if (Insurance.isEmpty()) {
						System.out.print("No Insurance selected");
					} else {
						driver.findElement(By.xpath("//input[@value='Modify INS/CP']")).click();
						String[] insVal = rwe.getCellData("Payless_GUI", k, 31).split("-");
						for (String e : insVal) {
							WebDriverWait wait1 = new WebDriverWait(driver, 10);
							if (e.equalsIgnoreCase("LDW")) {
								WebElement insurace1 = driver.findElement(By.id("menulist:rateshopContainer:resForm:coverageLdwYesNo"));
								Select insLDW = new Select(insurace1);
								wait1.until(ExpectedConditions.visibilityOf(insurace1));
								if (insurace1.isDisplayed()) {
									insLDW.selectByVisibleText("Yes");
								} else {
									break;
								}

							} else if (e.equalsIgnoreCase("PAI")) {
								WebElement insurace2 = driver.findElement(By.id("menulist:rateshopContainer:resForm:coveragePaiYesNo"));
								Select insPAI = new Select(insurace2);
								wait1.until(ExpectedConditions.visibilityOf(insurace2));
								if (insurace2.isDisplayed()) {
									insPAI.selectByVisibleText("Yes");
								} else {
									break;
								}
							} else if (e.equalsIgnoreCase("PEP")) {
								WebElement insurace3 = driver.findElement(By.id("menulist:rateshopContainer:resForm:coveragePepYesNo"));
								Select insPEP = new Select(insurace3);
								wait1.until(ExpectedConditions.visibilityOf(insurace3));
								if (insurace3.isDisplayed()) {
									insPEP.selectByVisibleText("Yes");
								} else {
									break;
								}

							} else if (e.equalsIgnoreCase("ALI")) {
								WebElement insurace4 = driver.findElement(By.id("menulist:rateshopContainer:resForm:coverageAliYesNo"));
								Select insALI = new Select(insurace4);
								wait1.until(ExpectedConditions.visibilityOf(insurace4));
								if (insurace4.isDisplayed()) {
									insALI.selectByVisibleText("Yes");
								} else {
									break;
								}
							} else if (e.equalsIgnoreCase("FSO")) {
								WebElement insurace5 = driver.findElement(By.id("menulist:rateshopContainer:resForm:fuelServiceOption"));
								Select insFSO = new Select(insurace5);
								wait1.until(ExpectedConditions.visibilityOf(insurace5));
								if (insurace5.isDisplayed()) {
									insFSO.selectByVisibleText("Yes");
								} else {
									break;
								}
							}

							else {
								break;
							}
						}

					}
				//Clicking on the update reservation button
				Thread.sleep(5000);
				driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
				Thread.sleep(3000);
				String Modmsg= driver.findElement(By.xpath("//form[@id='templateFatalForm']//div[@class='modal-body']")).getText();
				//System.out.println(Modmsg);
				rwe.setCellData("Payless_GUI", k, 46, Modmsg); // write respopup in excel
				if(Modmsg.contains("ERROR CODE"))
				{
				driver.findElement(By.xpath("//button[@id='templateFatalForm:templateFatalButton']")).click();
				}
				Thread.sleep(2000);
				
				//If we get any error message like checkout date/time cannot be before 
				String Errormsg = driver.findElement(By.xpath("//div[@id='rateshopErrorBarRes']//div[@class='modal-body']")).getText();
				rwe.setCellData("Payless_GUI", k, 46, Errormsg);
				Thread.sleep(2000);
				if(Errormsg.contains("date/time"))
				{
				driver.findElement(By.xpath("//div[@class='modal-body']//span[@class='ng-binding'][contains(text(),'Close')]")).click();
				}
				Thread.sleep(2000);
				
				//If we get valid output 
				String validmsg= driver.findElement(By.xpath("//form[@id='templateInfoForm']//div[@class='modal-body']")).getText();
				if (validmsg.contains("Reservation update completed successfully"))
				{
				driver.findElement(By.xpath("//button[@id='templateInfoForm:templateInfoButton']")).click();
				//Quick View
				int a1 =33;
				WebElement res = driver.findElement(By.cssSelector("#quickViewPanel > div.panel-body > table > tbody > tr:nth-child(8) > td > div:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > a > span"));

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
                                    	 if(a1<=45)
                                    	 {
                                            val = val.replaceAll("[*]", ""); // Remover '*' before rates
                                            rwe.setCellData("Payless_GUI", k, a1, val);
                                            a1++;
                                    	 }
                                     }
                              }
                       }
                }
				}
				
				//taking screenshot
				String ScreenShotPath = "C:\\Users\\cmn\\Desktop\\Selenium\\Screenshots\\Payless_GUI_Modify_Reservation";
				functions.ScreenCapture(ScreenShotPath , testcasename);
				

				/*
				 * Log out and close tabs
				 */

				functions.logout();
				Thread.sleep(1000);
				functions.closeWindows();
				test = extent.createTest(testcasename);
				if (rwe.getCellData("Payless_GUI", k, 33).isEmpty()) {
					rwe.setCellData("Payless_GUI", k, 47, "FAIL");
					test.log(Status.FAIL, "Fail");
				} else {
					rwe.setCellData("Payless_GUI", k, 47, "PASS");
					test.log(Status.PASS, "Pass");
				}
			}
			else
			{
				System.out.println("Execution status is N for iteration " + k + "...");
			}
		}
		
	} finally {
			// TODO: handle finally clause
		extent.flush();
	}
}
	
}
	
	

				
				
			