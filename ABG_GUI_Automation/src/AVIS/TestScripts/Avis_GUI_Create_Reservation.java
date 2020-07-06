 /**
 * 
 */
package AVIS.TestScripts;

import java.io.*;
import java.util.ArrayList;
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

/**
 * '#############################################################################################################################
 * '## SCRIPT NAME: GUI_CreateReservation_AVIS '## BRAND: AVIS '## DESCRIPTION:
 * Create Reservation for daily LOR with different insurance and counter
 * products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
 * All the required Test Data should be available in Test Data Sheet. '##
 * OUTPUT: Reservation should be created successfully.
 * 
 * 
 * HISTORY 05-SEP-2018 - GUIFunctions class created for GUI Common
 * functionalities and CR functionality
 * 
 * RS error handled
 * '#############################################################################################################################
 **/

public class Avis_GUI_Create_Reservation
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
	}
	

	@Test
	public void test() throws Exception
	
	{
		try {
			Properties prop = new Properties();
			FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\Downloads\\ABG-master\\ABG-master\\src\\AVIS\\TestData\\TestDataABGGUI.properties");
			prop.load(fis);
			//WebDriver driver;
			ChromeDriver driver = new ChromeDriver();
			GUIFunctions functions = new GUIFunctions(driver);
			System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
			driver.navigate().to(prop.getProperty("AvisURL"));
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			Thread.sleep(2000);
			functions.txt_userid.sendKeys(prop.getProperty("USERID"));
			Thread.sleep(500);
			functions.txt_password.sendKeys(prop.getProperty("PASSWORD"));
			Thread.sleep(500);
			functions.btn_login.click();
				/* Login */
				
				Thread.sleep(3000);
			// Read input from excel
			for (int k = 1; k <= 60; k++)
			{
				Avis_GUI_Create_Reservation avis = new Avis_GUI_Create_Reservation();
				
				ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Avis\\AVIS_GUITestData_CreateReservation.xlsx");
				String Execute = rwe.getCellData("Avis_GUI", k, 2);
				
				//********Delete the files in the folder********//
				File file = new File(prop.getProperty("ScreenshotAvis"));  

				String[] myFiles;    
				if (file.isDirectory()) {
				    myFiles = file.list();
				    for (int i = 0; i < myFiles.length; i++) {
				        File myFile = new File(file, myFiles[i]); 
				        myFile.delete();
				    }
				    int a = 28;
					
					System.out.println(" iteration " + k);
					String TCName = rwe.getCellData("Avis_GUI", k, 4);
					String clientURL = rwe.getCellData("Avis_GUI", k, 6);
					String outSTA = rwe.getCellData("Avis_GUI", k, 7);
					String thinClient = clientURL + outSTA;
					//String thinClient = clientURL;
					String uName = rwe.getCellData("Avis_GUI", k, 8);
					String pswd = rwe.getCellData("Avis_GUI", k, 9);
					String lstname = rwe.getCellData("Avis_GUI", k, 10);
					String fstname = rwe.getCellData("Avis_GUI", k, 11);
					String codte = rwe.getCellData("Avis_GUI", k, 12);
					String cotme = rwe.getCellData("Avis_GUI", k, 13);
					String insta = rwe.getCellData("Avis_GUI", k, 14);
					String cidte = rwe.getCellData("Avis_GUI", k, 15);
					String citme = rwe.getCellData("Avis_GUI", k, 16);
					String carGrp = rwe.getCellData("Avis_GUI", k, 17);
					String awd = rwe.getCellData("Avis_GUI", k, 18);
					String FTN = rwe.getCellData("Avis_GUI", k, 19);
					String cardname = rwe.getCellData("Avis_GUI", k, 20);
					String cardNo = rwe.getCellData("Avis_GUI", k, 21);
					// System.out.print("excel card number in script :"
					// +cardNo);
					String expireMonth = rwe.getCellData("Avis_GUI", k, 22);
					String expireYear = rwe.getCellData("Avis_GUI", k, 23);
					String reason = rwe.getCellData("Avis_GUI", k, 24);

				    
				//Launch the application and Enter the username and password
				  
					//driver.navigate().to(prop.getProperty("URL"));
				    
				if (Execute.equals("Y"))
				{
					
					
					String ScreenshotPath = prop.getProperty("ScreenshotAvis");	
					
					/* Open GUI URL's */
					// System.out.println(" token URL value : " + tokenURL);
					//*******Screenshot path and test name*********//
					
					String testcasename = TCName;
					String xfilepath = prop.getProperty("ExcelPathAvis") +testcasename+ ".xlsx";
					test = extent.createTest(TCName);
					
					functions.openURL(thinClient);
					/* Login */
					//functions.login(uName, pswd);
					functions.navigateToTab("ReservationRates");
					Thread.sleep(2000);				
					
					/* Enter Customer Informations */
					/* Enter FTN */
					if (rwe.getCellData("Avis_GUI", k, 19).isEmpty()) {
						System.out.println("No FTN added");
					} else {
						functions.expandToggleBtn();
						Thread.sleep(2000);
						functions.enterFTN(FTN);
					}
					//driver.navigate().refresh();
					Thread.sleep(2000);
					//functions.navigateToTab("ReservationRates");
					//driver.switchTo().alert().accept();
					//functions.enterCustomerName(lastname,firstname);
					Thread.sleep(3000);
					functions.enterCustomerName(lstname, fstname);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//functions.enterDriverDetail(drCountry,drState,drNumber, drDOB, drCompany,addr1,addr2, addr3, contact);
					Thread.sleep(2000);
					driver.findElement(By.id("menulist:rateshopContainer:resForm:pickupStation:pickupStation")).clear();
					Thread.sleep(2000);
					driver.findElement(By.id("menulist:rateshopContainer:resForm:pickupStation:pickupStation")).sendKeys(outSTA);
					functions.enterCustomerInformation(codte, cotme, insta, cidte, citme);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//enterCustomerInformation(lstname,fstname,codte,cotme,insta,cidte,citme);
					/* Enter AWD */
					if (awd.isEmpty()) {
						System.out.println("No Avis Discount Number Added");
					} else {
						functions.enterAWD(awd);
					}

					/* Select car group */
					functions.selectCarGroupByVT(carGrp);
					Thread.sleep(2000);

					/* RATE SHOP */
					avis.clickRateshopSearchBtn(driver);
					ArrayList<WebElement> radio = (ArrayList<WebElement>) driver
							.findElements(By.xpath("//input[@name='radioRate'and @type='radio']"));
					for (int i = 0; i < radio.size(); i++) {
						if ((radio.get(i).isDisplayed()) && (radio.get(i).isEnabled())) {
							radio.get(i).click();
							if (radio.get(i).isSelected()) {
								functions.clickSelectRateBtn();
								/* Enter MOP details */
								functions.expandPaymentInfoSection();
								functions.enterPaymentInformations(cardname, cardNo, expireMonth, expireYear, reason);

								/* Add Insurances */
								Thread.sleep(5000);
								functions.expandProtectionCoverageSection();
								if (rwe.getCellData("Avis_GUI", k, 25).isEmpty()) {
									System.out.print("No Insurance selected");
								} else {
									String[] insVal = rwe.getCellData("Avis_GUI", k, 25).split("-");
									for (String e : insVal) {
										WebDriverWait wait1 = new WebDriverWait(driver, 10);
										if (e.equalsIgnoreCase("LDW")) {
											WebElement insurace1 = driver.findElement(
													By.id("menulist:rateshopContainer:resForm:coverageLdwYesNo"));
											Select insLDW = new Select(insurace1);
											wait1.until(ExpectedConditions.visibilityOf(insurace1));
											if (insurace1.isDisplayed()) {
												insLDW.selectByVisibleText("Yes");
											} else {
												break;
											}
										} else if (e.equalsIgnoreCase("PAI")) {
											WebElement insurace2 = driver.findElement(
													By.id("menulist:rateshopContainer:resForm:coveragePaiYesNo"));
											Select insPAI = new Select(insurace2);
											wait1.until(ExpectedConditions.visibilityOf(insurace2));
											if (insurace2.isDisplayed()) {
												insPAI.selectByVisibleText("Yes");
											} else {
												break;
											}
										} else if (e.equalsIgnoreCase("PEP")) {
											WebElement insurace3 = driver.findElement(
													By.id("menulist:rateshopContainer:resForm:coveragePepYesNo"));
											Select insPEP = new Select(insurace3);
											wait1.until(ExpectedConditions.visibilityOf(insurace3));
											if (insurace3.isDisplayed()) {
												insPEP.selectByVisibleText("Yes");
											} else {
												break;
											}
										} else if (e.equalsIgnoreCase("ALI")) {
											WebElement insurace4 = driver.findElement(
													By.id("menulist:rateshopContainer:resForm:coverageAliYesNo"));
											Select insALI = new Select(insurace4);
											wait1.until(ExpectedConditions.visibilityOf(insurace4));
											if (insurace4.isDisplayed()) {
												insALI.selectByVisibleText("Yes");
											} else {
												break;
											}
										} else if (e.equalsIgnoreCase("FSO")) {
											WebElement insurace5 = driver.findElement(
													By.id("menulist:rateshopContainer:resForm:fuelServiceOption"));
											Select insFSO = new Select(insurace5);
											wait1.until(ExpectedConditions.visibilityOf(insurace5));
											if (insurace5.isDisplayed()) {
												insFSO.selectByVisibleText("Yes");
											} else {
												break;
											}
										} else {
											break;
										}
									}
								}

								/*
								 * Add CounterProducts
								 */
								if (rwe.getCellData("Avis_GUI", k, 26).isEmpty()) {
									System.out.print("No CounterProduct selected");
								} else {
									String[] cpVal = rwe.getCellData("Avis_GUI", k, 26).split("-");
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
												WebElement cp5 = driver.findElement(By.id("productQuantityYesNo11"));
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

								/*
								 * Create reservation
								 */
								functions.clickCreateReservationBtn();
								functions.ScreenCapturedate(ScreenshotPath,TCName);
								Thread.sleep(1000);
								String Resmsg = driver
										.findElement(By.xpath("//*[@id='templateInfoForm:templateInfoMsg']")).getText();
								functions.ScreenCapturedate(ScreenshotPath,TCName);
								rwe.setCellData("Avis_GUI", k, 40, Resmsg); // write
																			// respopup
																			// in
								Thread.sleep(2000);											// excel
								String resno = Resmsg.substring(54,65);
								rwe.setCellData("Avis_GUI", k, 27, resno);
								Thread.sleep(1000);
								driver.findElement(By.xpath("//*[@id='templateInfoForm:templateInfoButton']")).click(); // clicks OK button in Res popup
								Thread.sleep(5000);					  
								functions.ScreenCapturedate(ScreenshotPath, TCName);

								//String resno= driver.findElement(By.xpath("//*[@class='qv-link qvResResNo']")).getText();
								//System.out.println(resno);
								
								/*
								 * to print QV data in Excel
								 */
								//WebElement res = driver.findElement(By.cssSelector(
										///"#quickViewPanel > div.panel-body > table > tbody > tr:nth-child(8) > td > div:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > a > span"));
								WebElement res = driver.findElement(By.xpath("//div[@id='QuickView-qvRes']//span[@class='ng-binding'][contains(text(),'Quick View')]"));
								//functions.ScreenCapturedate(ScreenshotPath,TCName);
								try {
									if (res.getText().isEmpty()) {
										WebDriverWait wait1 = new WebDriverWait(driver, 20);
										wait1.until(ExpectedConditions.visibilityOf(res));
									}
								} catch (Exception e2) {
									e2.printStackTrace();
								}
								WebElement table = driver.findElement(By.id("resQvForm"));
								ArrayList<WebElement> rows = (ArrayList<WebElement>) table
										.findElements(By.tagName("tr"));
								for (int i1 = 1; i1 < rows.size(); i1++) {
									ArrayList<WebElement> cells = (ArrayList<WebElement>) rows.get(i1)
											.findElements(By.tagName("td"));
									for (int j = 0; j < cells.size(); j++) {
										String val = cells.get(j).getText();
										if (val.isEmpty()) {
											break;
										} else {
											if (j == 2) {
												val = val.replaceAll("[*]", ""); // Remover * before rates
												rwe.setCellData("Avis_GUI", k, a, val);
												a++;
											}
										}
									}
								}
								/*
								 * Log out and close tabs
								 */
								//functions.ScreenCapturedate(ScreenshotPath,TCName);
								//functions.logout();
								//Thread.sleep(1000);
								//functions.closeWindows();

								//test = extent.createTest(TCName);

								if (rwe.getCellData("Avis_GUI", k, 27).isEmpty()) {
									rwe.setCellData("Avis_GUI", k, 41, "FAIL");
									test.log(Status.FAIL, "Fail");
								} else {
									rwe.setCellData("Avis_GUI", k, 41, "PASS");
									test.log(Status.PASS, "Pass");
								}
							} // end of inner if
						}  // end of if
						else // Select available base rate from Standard rate
								// table
						{
							System.out.print(" No Base Rates are Available ");
							String RateShopErrorMsg = driver
									.findElement(
											By.xpath("//*[@id='resRateLookupDlg:rateLookupForm:counterRates_data']"))
									.getText();
							String[] RateShopErrorMsg1 = RateShopErrorMsg.split("Base Rate");
							String RSerrmsg = RateShopErrorMsg1[1];
							System.out.println("Rate shop error msg:" + RSerrmsg);
							rwe.setCellData("Avis_GUI", k, 40, RSerrmsg);
							functions.ScreenCapturedate(ScreenshotPath,TCName);
							driver.findElement(By.xpath("//input[@id='cancelButton']")).click();
							Thread.sleep(2000);
							functions.openURL(thinClient);
							/* Login */
							//functions.login(uName, pswd);
							functions.navigateToTab("ReservationRates");
							Thread.sleep(2000);				
							//functions.logout();
							//functions.closeWindows();
						} // end of else
						
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
						       File f = new File(prop.getProperty("ScreenshotAvis"));

						       // Populates the array with names of files and directories
						       pathnames = f.list();

						       // For each pathname in the pathnames array
						       for (String pathname : pathnames) {
						             // Print the names of files and directories

						             InputStream inputStream = new FileInputStream(
						            		 prop.getProperty("ScreenshotAvis")+pathname);
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
					} 
				} // end of Iteration if

				else {
					System.out.println("Execution status is N for iteration " + k + "...");
				}

			}
		
	}// end of for
			functions.logout();
			Thread.sleep(1000);
			functions.closeWindows();
} finally{
	extent.flush();
}
}
}