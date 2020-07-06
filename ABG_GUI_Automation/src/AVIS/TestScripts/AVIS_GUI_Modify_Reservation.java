package AVIS.TestScripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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


public class AVIS_GUI_Modify_Reservation
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
		// Read input from excel
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
		for (int k = 1; k <= 60; k++)
		{
			AVIS_GUI_Modify_Reservation avis = new AVIS_GUI_Modify_Reservation();
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Avis\\AVIS_GUITestData_ModifyReservation.xlsx");
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
			}
			    
			    int a = 27;
				
				System.out.println(" iteration " + k);
				String TCName    = rwe.getCellData("Avis_GUI", k, 4);
				//String tokenURL        = rwe.getCellData("Avis_GUI", k, 6);
				String clientURL       = rwe.getCellData("Avis_GUI", k, 6);
				String outSTA          = rwe.getCellData("Avis_GUI", k, 7);
				String thinClient = clientURL + outSTA;
				String uName           = rwe.getCellData("Avis_GUI", k, 8);
				String pswd            = rwe.getCellData("Avis_GUI", k, 9);
				String lstname         = rwe.getCellData("Avis_GUI", k, 10);
				String fstname         = rwe.getCellData("Avis_GUI", k, 11);
				String comonth		   = rwe.getCellData("Avis_GUI", k, 12);	
				String codte           = rwe.getCellData("Avis_GUI", k, 13);
				String cotme           = rwe.getCellData("Avis_GUI", k, 14);
				String insta           = rwe.getCellData("Avis_GUI", k, 15);
				String cimonth         = rwe.getCellData("Avis_GUI", k, 16);
				String cidte           = rwe.getCellData("Avis_GUI", k, 17);
				String citme           = rwe.getCellData("Avis_GUI", k, 18);
				String carGrp          = rwe.getCellData("Avis_GUI", k, 19);
				String awd             = rwe.getCellData("Avis_GUI", k, 20);
				String creditcard        = rwe.getCellData("Avis_GUI", k, 21);
				String cardname        = rwe.getCellData("Avis_GUI", k, 22);
				String cardNumber      = rwe.getCellData("Avis_GUI", k, 23);
				String expireMonth     = rwe.getCellData("Avis_GUI", k, 24);
				String expireYear      = rwe.getCellData("Avis_GUI", k, 25);
				String Ftncode         = rwe.getCellData("Avis_GUI", k, 26);
				String Ftnno           = rwe.getCellData("Avis_GUI", k, 27);
				String Reservation_No  = rwe.getCellData("Avis_GUI", k, 28);
				String Modout          = rwe.getCellData("Avis_GUI", k, 29);
				String Insurance       = rwe.getCellData("Avis_GUI", k, 30);
				String Counterproducts = rwe.getCellData("Avis_GUI", k, 31);
			
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
				
				//Entering the reservation number in search field
				Thread.sleep(2000);
				//functions.enterreseinsearchfield(Reservation_No);
				functions.btn_ressearch.click();
				Thread.sleep(2000);
				functions.txt_displres.click();
				Thread.sleep(2000);
				functions.txt_displres.sendKeys(Reservation_No);
				Thread.sleep(2000);
				functions.btn_popupsearch.click();
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
					Thread.sleep(2000);
					driver.findElement(By.cssSelector("input[id='menulist:rateshopContainer:resForm:firstName']")).click();
					driver.findElement(By.cssSelector("input[id='menulist:rateshopContainer:resForm:firstName']")).clear();
					Thread.sleep(1000);
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
						Thread.sleep(2000);
						driver.findElement(By.xpath("//input[@id= 'menulist:rateshopContainer:resForm:checkout1_hid']")).click();
						driver.findElement(By.xpath("//input[@id= 'menulist:rateshopContainer:resForm:checkout1_hid']")).clear();
						driver.findElement(By.xpath("//input[@id= 'menulist:rateshopContainer:resForm:checkout1_hid']")).sendKeys(comonth);
						Thread.sleep(2000);
						driver.findElement(By.xpath("//div[@class= 'col-md-2 col-sm-6 div-zero-padding-left-0 margin-top-35']")).click();
						
						/*while(!driver.findElement(By.cssSelector("[class='ui-datepicker-title'] [class='ui-datepicker-month']")).getText().contains(comonth))
						{
							driver.findElement(By.cssSelector("[class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all'] [class='ui-icon ui-icon-circle-triangle-e']")).click();
						}
						Thread.sleep(2000);
						List<WebElement> dates = driver.findElements(By.xpath("//*[@data-handler='selectDay']"));
						
						int Count = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).size();
						Thread.sleep(2000);
						for(int i=0;i<Count;i++)
						{
							String text = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).getText();
							if(text.equalsIgnoreCase(codte))
							{
								driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).click();
								break;
							}
						}*/
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
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin1_hid']")).sendKeys(cimonth);
						Thread.sleep(2000);
						driver.findElement(By.xpath("//div[@class= 'col-md-2 col-sm-6 div-zero-padding-left-0 margin-top-35']")).click();
						//Picking checkin date and time
						/*while(!driver.findElement(By.cssSelector("[class='ui-datepicker-title'] [class='ui-datepicker-month']")).getText().contains(cimonth))
						{
							driver.findElement(By.cssSelector("[class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all'] [class='ui-icon ui-icon-circle-triangle-e']")).click();
						}
						Thread.sleep(2000);
						List<WebElement> datescheckin = driver.findElements(By.xpath("//*[@data-handler='selectDay']"));
						Thread.sleep(2000);
						int Count1 = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).size();
						Thread.sleep(2000);
						for(int i=0;i<Count1;i++)
						{
							String text = driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).getText();
							if(text.equalsIgnoreCase(cidte))
							{
								driver.findElements(By.xpath("//*[@data-handler='selectDay']")).get(i).click();
								break;
							}
						}*/
						//driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin1_hid']")).sendKeys(cidte);
						Thread.sleep(3000);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).click();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).clear();
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).sendKeys(citme);
						driver.findElement(By.xpath("//input[@id='menulist:rateshopContainer:resForm:checkin2']")).sendKeys(Keys.TAB);
						Thread.sleep(3000);
						//driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
						//Thread.sleep(3000);
						
						//Entering the Rate code after change in checkout time
						driver.findElement(By.xpath("//input[@ng-model='resMB.res.rateCode']")).click();
						driver.findElement(By.xpath("//input[@ng-model='resMB.res.rateCode']")).clear();
						avis.clickRateshopSearchBtn(driver);
						ArrayList<WebElement> radio = (ArrayList<WebElement>) driver
								.findElements(By.xpath("//input[@name='radioRate'and @type='radio']"));
						for (int i = 0; i < radio.size(); i++) {
							if ((radio.get(i).isDisplayed()) && (radio.get(i).isEnabled())) {
								radio.get(i).click();
								if (radio.get(i).isSelected()) {
									functions.clickSelectRateBtn();
								}
							}
						}
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
				
				//Modifying the counterproduct
					if(Counterproducts.isEmpty()) {
						System.out.print("No CounterProduct selected");
						// break;
					} else {
						String[] cpVal = rwe.getCellData("Avis_GUI", k, 31).split("-");
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
				
				
				//Clicking on the update reservation button
					Thread.sleep(3000);
				driver.findElement(By.xpath("//input[@id='footerForm:footerUpdateRes']")).click();
				Thread.sleep(3000);
				/*driver.switchTo().alert().accept();
				Thread.sleep(1000);
				driver.switchTo().alert().accept();
				Thread.sleep(1000);
				driver.switchTo().alert().accept();
				Thread.sleep(1000);
				driver.switchTo().alert().accept();*/
				Thread.sleep(1000);
				String Modmsg= driver.findElement(By.xpath("//form[@id='templateFatalForm']//div[@class='modal-body']")).getText();
				System.out.println(Modmsg);
				//String Modmsg = driver.findElement(By.xpath("*//[@id='templateInfoForm:templateInfoMsg']")).getText();
				//System.out.println(Modmsg);
				//rwe.setCellData("Avis_GUI", k, 45, Modmsg); // write respopup in excel
				if(Modmsg.contains("ERROR CODE"))
				{
				rwe.setCellData("Avis_GUI", k, 45, Modmsg); 
				driver.findElement(By.xpath("//button[@id='templateFatalForm:templateFatalButton']")).click();
				
				}
				Thread.sleep(2000);
				
				//If we get any error message like checkout date/time cannot be before 
				String Errormsg = driver.findElement(By.xpath("//div[@id='rateshopErrorBarRes']//div[@class='modal-body']")).getText();
				//String Errormsg = driver.findElement(By.xpath("*//[@id='templateInfoForm:templateInfoMsg']")).getText();
				Thread.sleep(2000);
				if(Errormsg.contains("date/time"))
				{
				rwe.setCellData("Avis_GUI", k, 45, Errormsg);
				driver.findElement(By.xpath("//div[@class='modal-body']//span[@class='ng-binding'][contains(text(),'Close')]")).click();
				}
				Thread.sleep(2000);
				
				//If we get valid output 
				//String validmsg= driver.findElement(By.xpath("*//[@id='templateInfoForm:templateInfoMsg']")).getText();
				//String validmsg= driver.findElement(By.xpath("//form[@id='templateFatalForm']//div[@class='modal-body']")).getText();
				
					String Resmsg = driver
							.findElement(By.xpath("//form[@id='templateFatalForm']//div[@class='modal-body']")).getText();
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					rwe.setCellData("Avis_GUI", k, 45, Resmsg); // write
																// respopup
																// in
					Thread.sleep(2000);											// excel
					String resno1 = Resmsg.substring(43,54);
					rwe.setCellData("Avis_GUI", k, 32, resno1);
					Thread.sleep(1000);
					functions.ScreenCapturedate(ScreenshotPath, TCName);
				driver.findElement(By.xpath("//button[@id='templateInfoForm:templateInfoButton']")).click();
				Thread.sleep(3000);                                    
				//Quick View
				int a1 =33;
				//WebElement res = driver.findElement(By.cssSelector("#quickViewPanel > div.panel-body > table > tbody > tr:nth-child(8) > td > div:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > a > span"));
				WebElement res = driver.findElement(By.xpath("//div[@id='QuickView-qvRes']//span[@class='ng-binding'][contains(text(),'Quick View')]"));
				
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
                                    	 if(a1<=44)
                                    	 {
                                            val = val.replaceAll("[*]", ""); // Remover '*' before rates
                                            rwe.setCellData("Avis_GUI", k, a1, val);
                                            a1++;
                                    	 }
                                     }
                              }
                       }
                }
				
				
				test = extent.createTest(TCName);
				if (rwe.getCellData("Avis_GUI", k, 32).isEmpty())
				{
					rwe.setCellData("Avis_GUI", k, 46, "FAIL");
					test.log(Status.FAIL, "Fail");
				}
				else
				{
					rwe.setCellData("Avis_GUI", k, 46, "PASS");
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
			
				}//end of if statement
	
		/*
		 * Log out and close tabs
		 */
		}//end of for statement
		functions.logout();
		Thread.sleep(1000);
		functions.closeWindows();
		
		}
		finally
		{
			// TODO: handle finally clause
			extent.flush();
		}
				}
				
			}
		
	

				
				
			