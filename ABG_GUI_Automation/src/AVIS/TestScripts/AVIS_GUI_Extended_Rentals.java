package AVIS.TestScripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Hashtable;
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
 * '## SCRIPT NAME: AVIS_GUI_Extended_Rentals '## BRAND: AVIS '## DESCRIPTION:
 * Creating RA numbers for Extended Rentals '## FUNCTIONAL AREA :
 * Checkout,Checkin, Display Rental screen '## PRECONDITION:All the required
 * Test Data should be available in Test Data Sheet. '##OUTPUT: Extended Rentals
 * should be created successfully.
 * 
 * HISTORY 12-SEP-2018 - GUIFunctions class created for GUI Common
 * functionalities and CR functionality
 * '#############################################################################################################################
 **/

public class AVIS_GUI_Extended_Rentals {

	public void clickRateshopSearchBtn(ChromeDriver driver) {
		JavascriptExecutor jse = (JavascriptExecutor) driver;
		String clickSearchJS = "document.getElementById('searchCommandLinkResRateCode').click()";
		jse.executeScript(clickSearchJS);
	}

	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport() {

		extent = Extentmanager.GetExtent();
		// test = extent.createTest("GUI");

	}

	@SuppressWarnings("unlikely-arg-type")
	// public static void main(String[] args) throws IOException, Exception,
	// FileNotFoundException {

	@Test
	public void test() throws Exception {
		try {

			// Read input from excel

			for (int k = 1; k <= 100; k++)
			{
				AVIS_GUI_Extended_Rentals avis = new AVIS_GUI_Extended_Rentals();
				ReadWriteExcel rwe = new ReadWriteExcel(
						"C:\\Users\\cmn\\Desktop\\Selenium\\TestData\\AVIS_GUITestData_Extended_Rental.xlsx");
				String Execute = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 2);

				if (Execute.equals("Y")) {
					System.setProperty("webdriver.chrome.driver",
							"C:\\Users\\cmn\\Desktop\\Selenium\\chromedriver.exe");
					ChromeDriver driver = new ChromeDriver();
					GUIFunctions functions = new GUIFunctions(driver);
					int a = 27;
					// ChromeDriver driver = new ChromeDriver();
					// GUIWebDriverFunctions wdfunctions = new GUIWebDriverFunctions(driver);
					driver.manage().window().maximize();
					driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
					System.out.println(" iteration " + k);
					String testcasename = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 4);
					String tokenURL = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 6);
					String clientURL = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 7);
					String outSTA = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 8);
					String thinClient = clientURL + outSTA;
					String uName = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 9);
					String pswd = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 10);
					String Resno = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 11);
					String Coemail = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 12);
					String COMILEAGE1 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 13);
					String COMVA = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 14);
					String RANumber_1 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 15);
					String RANumber_2 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 16);
					String MilageIn2 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 17);
					String MilageIn3 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 18);
					String Action_code_CI = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 19);

					/* Open GUI URL's */
					driver.get(thinClient);

					/* Login */
					functions.login(uName, pswd);
					driver.navigate().refresh();
					functions.navigateToTab("CheckOut");
					Thread.sleep(8000L);

					// Step1: On checkout screen--> click delay toggle and enter Res no in the field
					// and search
					driver.findElement(By.xpath("//span[@id='checkoutToggle']")).isDisplayed();//gui 2 link
					driver.findElement(By.xpath("//span[@id='checkoutToggle']")).click();//gui 2 link
					//driver.findElement(By.id("delayBtn")).isDisplayed();//gui 1 link
					//driver.findElement(By.id("delayBtn")).click();//gui 1 link
					Thread.sleep(3000);
					
					// Enter the res no in search field and click on search button
					driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).click();
					driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).sendKeys(Resno);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@ng-click='directSearch()']")).click();
					Thread.sleep(15000);

					// Step2: Enter the soft prompt(email address)and click on continue
					String email = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut"))
							.getText();
					if (email.contains("PLEASE INPUT CUSTOMER EMAIL ADDRESS")) {
						// boolean emailfield =
						// driver.findElement(By.xpath("//input[@class='form-control ng-pristine
						// ng-valid allowBackspace ng-touched']")).isDisplayed();
						// System.out.println(emailfield);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:wizconEmail")).click();
						Thread.sleep(4000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:wizconEmail")).sendKeys(Coemail);
						Thread.sleep(2000);
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:wizconSubmitButton']"))
								.isDisplayed();
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:wizconSubmitButton']"))
								.click();
						Thread.sleep(15000);
					}

					// If DOB is not entered
					String DOB = driver
							.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']"))
							.getAttribute("value");
					System.out.println("The Date od Birth entered is " + DOB);
					if (DOB.isEmpty()) {
						driver.findElement(
								By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']"))
								.sendKeys("06/09/77");
						System.out.println("Entered the Date of Birth");
					}

					// Step3:Click on delay continue button on checkout screen
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).click();
					Thread.sleep(8000);

					// Step4:Click on Ok Button on the pop up window displayed
					String partialcheckout = driver.switchTo().alert().getText();
					if (partialcheckout.contains("Are you sure you want to do a partial checkout?")) {
						driver.switchTo().alert().accept();
						Thread.sleep(5000);
					}

					Thread.sleep(9000);
					// Enter MVA number
					String comilenmva = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut"))
							.getText();
					if (comilenmva.contains("ERROR - SEE HIGHLIGHTED FIELDS")) {
						// Enter all the required fields checkout miles and MVA number
						Thread.sleep(5000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.isDisplayed();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.click();
						Thread.sleep(5000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.sendKeys(COMILEAGE1);
						Thread.sleep(2000);
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.isDisplayed();
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.click();
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.sendKeys(COMVA);
						Thread.sleep(2000);
						// Click on Continue Button
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']"))
								.isDisplayed();
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']"))
								.click();
						Thread.sleep(8000);
					}

					// Enter the valid RA number in the field
					String RANumber1 = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
					if (RANumber1.contains("NON NUMERIC RA NUMBER")) {
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(RANumber_1);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
						Thread.sleep(5000);
					}

					String documentnoinuse = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
					if (documentnoinuse.contains("DOCUMENT NUMBER ALREADY IN USE")) {
						while(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText().contains("C/O CAR INFO DISCREPANCIES") == false) {

							
							//String new_Ra2 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
							int new_RA = get_RA_Number(Integer.parseInt(RANumber_1));
			
								driver.findElement(
										By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
								driver.findElement(
										By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
										.sendKeys(Integer.toString(new_RA));
								driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton"))
										.click();
								Thread.sleep(5000);

							}
					
						} 
					

					// For message as C/O CAR INFO DISCREPANCIES
					String afterRa1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText();
					if (afterRa1.contains("C/O CAR INFO DISCREPANCIES")) or (afterRa1.contains("DO NOT RENT THIS CAR"));
					{
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(MilageIn2);
						Thread.sleep(2000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue")).sendKeys(COMVA);
						Thread.sleep(2000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
						Thread.sleep(10000);
					}

					// Click on complete checkout button
					driver.findElement(By.id("footerForm:completeCheckoutButton")).isDisplayed();
					driver.findElement(By.id("footerForm:completeCheckoutButton")).click();
					Thread.sleep(10000);

					// Capture back the values into Xls
					String strCOLNFNGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 20, strCOLNFNGetText);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
					String strCOVehicleModelGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 21, strCOVehicleModelGetText);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
					String strCOResNoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 22, strCOResNoGetText);
					System.out.println(" Reservation No value is " + strCOResNoGetText);
					String strCOMVAGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 23, strCOMVAGetText);
					System.out.println(" MVA No value is " + strCOMVAGetText);
					String strCORANoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 24, strCORANoGetText);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println(" 1st RA Number value is  " + strCORANoGetText);
					String strCOLicensePlateNoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 25, strCOLicensePlateNoGetText);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
					String strCOEstimatedCostGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 26, strCOEstimatedCostGetText);
					System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
					String strCOSystemMsgGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 27, strCOSystemMsgGetText);
					System.out.println(" 1st CheckOut Complete System Message value is " + strCOSystemMsgGetText);
					Thread.sleep(7000);

					// Click on View Rental Agreement button
					System.out.println(driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).isDisplayed());
					
					driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).click();
					Thread.sleep(15000);

					// Click on Extended button in display Rental screen
					//driver.findElement(By.xpath("//form[@id='menulist:dispormodSubMenu']/div[8]/button/span")).isDisplayed();
					//driver.findElement(By.xpath("//form[@id='menulist:dispormodSubMenu']/div[8]/button/span")).click();
					driver.findElement(By.xpath("//button[@class='rentalExtendClass']")).isDisplayed();
					driver.findElement(By.xpath("//button[@class='rentalExtendClass']")).click();

					// Enter RAnumber2 and Mileagein in extended rental popup window
					String extendedrollover = driver.findElement(By.xpath("//div[@id='extendedrollover']/div/h1")).getText();
					if (extendedrollover.contains("Extended Rental")) {
						//int new_RA1 = get_RA_Number(Integer.parseInt(RANumber_2));
						driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).isDisplayed();
						//driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).click();
						Thread.sleep(2000);
						driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).sendKeys(RANumber_2);
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).isDisplayed();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).click();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).sendKeys(MilageIn3);
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedDone")).isDisplayed();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedDone")).click();
						Thread.sleep(9000);
						//Click on the "Ok" button "ERROR - SEE HIGHLIGHTED FIELDS : MILEAGE"
						driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).isDisplayed();
						driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).click();
						Thread.sleep(15000);
					}
					
					//Click on delayed toggle
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					//Select purchase fuel
					Select purchasefuel = new Select(driver.findElement(By.xpath("//select[@onchange='disableFuelLevel(this)']")));
					purchasefuel.selectByVisibleText("Yes");
					Thread.sleep(3000);
					
					//click on "complete checkin " button 
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).click();
					Thread.sleep(10000);
					
					//Enter the Second RA number and click on Roll over button
					driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid allowBackspace ng-touched']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid allowBackspace ng-touched']")).click();
					driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid allowBackspace ng-touched']")).sendKeys(RANumber_2);
				
					//Click on Roll over button
					driver.findElement(By.xpath("//button[@id='checkinDialogs:checkInSuccessForm:extndDlgRolloverButton']")).isDisplayed();
					driver.findElement(By.xpath("//button[@id='checkinDialogs:checkInSuccessForm:extndDlgRolloverButton']")).click();
					Thread.sleep(9000);

					// Capture back the values into Xls
					String strCOLNFNGetText1 = driver.findElement(By.xpath("//label[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteName']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 28, strCOLNFNGetText1);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText1);
					
					String strCOVehicleModelGetText1 = driver.findElement(By.xpath("//label[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteVehicle']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 29, strCOVehicleModelGetText1);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText1);
			
					String strCOMVAGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteMVA']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 31, strCOMVAGetText1);
					System.out.println(" MVA No value is " + strCOMVAGetText1);
					
					String strCORANoGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteRANumber']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 32, strCORANoGetText1);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println("2nd RA Number value is " + strCORANoGetText1);
					
					String strCOLicensePlateNoGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 33, strCOLicensePlateNoGetText1);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText1);
					
					String strCOEstimatedCostGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteEstimatedCost']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 34, strCOEstimatedCostGetText1);
					System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText1);
					
					String strCOSystemMsgGetText1 = driver.findElement(By.xpath("//textarea[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteOut']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 35, strCOSystemMsgGetText1);
					System.out.println(" 2nd CheckOut Complete System Message value is " + strCOSystemMsgGetText1);
					Thread.sleep(7000);
					
					//Click on Done button
					driver.findElement(By.xpath("//input[@id='checkinRolloverDialog:checkinRolloverCompleteForm:mleasecheckoutCompleteDoneButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRolloverDialog:checkinRolloverCompleteForm:mleasecheckoutCompleteDoneButton']")).click();
					Thread.sleep(8000);
					
					//Click on toggle button and Enter the RA_number2 in search field and click on search in checkin screen
					//driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					//driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).click();
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).sendKeys(RANumber_2);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='checkinSearchCommandLink']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinSearchCommandLink']")).click();
					Thread.sleep(10000);
					
					//Click on delayed toggle
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					//Select purchase fuel
					Select purchasefuel1 = new Select(driver.findElement(By.xpath("//select[@onchange='disableFuelLevel(this)']")));
					purchasefuel1.selectByVisibleText("Yes");
					Thread.sleep(3000);
					
					
					//click on "complete checkin " button 
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).click();
					Thread.sleep(10000);
					
					
					//Enter the mileage in pop up displayed and click on Continue button
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).click();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).clear();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).sendKeys(MilageIn3);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
					Thread.sleep(9000);
					
					//Click on Continue button again in next window
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
					Thread.sleep(9000);
					
					//Capture the checkin details back to xls
					String strCOLNFNGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successName']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 36, strCOLNFNGetText2);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText2);
					String strCOVehicleModelGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successVehicle']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 37, strCOVehicleModelGetText2);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText2);
					String strCORANoGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successRA']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 38, strCORANoGetText2);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println(" 2nd RA Number value is " + strCORANoGetText2);
					String strCOLicensePlateNoGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successPlate']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 39, strCOLicensePlateNoGetText2);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText2);
					String strCOSystemMsgGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successOutMsg']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 40, strCOSystemMsgGetText2);
					System.out.println(" 2nd Checkin Complete System Message value is " + strCOSystemMsgGetText2);
					Thread.sleep(7000);
					
					
					//Click on done button on the last screen or view rental
					//driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).click();
					driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:viewCheckInRental']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:viewCheckInRental']")).click();
					Thread.sleep(18000);
					
					// taking screenshot
					String ScreenShotPath = "C:\\Users\\cmn\\Desktop\\Selenium\\Screenshots\\Avis_GUI_Extended_Rentals\\";
					functions.ScreenCapture(ScreenShotPath, testcasename);
					

					/*
					 * Log out and close tabs
					 */

					functions.logout();
					Thread.sleep(1000);
					functions.closeWindows();
					test = extent.createTest(testcasename);
					if (rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 38).isEmpty()) {

						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 41, "FAIL");
						test.log(Status.FAIL, "Fail");
					} else {
						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 41, "PASS");
						test.log(Status.PASS, "Pass");
					}
				} else {
					System.out.println("Execution status is N for iteration " + k + "...");
				}

			}

		} finally {
			extent.flush();
			// TODO: handle finally clause
		}
	}

	private void or(boolean contains) {
		// TODO Auto-generated method stub
		
	}

	public static int get_RA_Number(int RA_Num) {

		if (Integer.toString((RA_Num)).endsWith("6")) {
			RA_Num = RA_Num + 4;
			System.out.println(RA_Num);
		}

		else {
			RA_Num = RA_Num + 11;
			//System.out.println(RA_Num);
		}
		return RA_Num;
		
		
	}
}
