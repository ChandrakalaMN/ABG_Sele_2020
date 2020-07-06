package Payless.TestScripts;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;

public class Payless_GUI_VoidRental
{
	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport() {

		extent = Extentmanager.GetExtent();
		//test = extent.createTest("GUI");

	}
	//public static void main(String[] args) throws IOException, Exception, FileNotFoundException, InterruptedException  
	//{
	@Test
	public void test() throws Exception {
		try {
		// Read input from excel
		for (int k = 1; k <= 20; k++)
		{
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\GUI_Automation\\Selenium\\GUI_TestDataSheets\\Payless_GUITestData_VoidRental.xlsx");
			String Execute = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 2);
			System.out.println("Iteration :" +k);
			
			if (Execute.equals("Y"))
			{
				System.setProperty("webdriver.chrome.driver", "C:\\Users\\csurek\\Videos\\chromedriver_win32 (1)\\chromedriver.exe");
				ChromeDriver driver = new ChromeDriver();
				GUIFunctions functions = new GUIFunctions(driver);
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				System.out.println(" iteration " + k);
				String testcasename   = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 5);
				String clientURL   = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 6);
				String outSTA      = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 7);
				String thinClient  = clientURL+outSTA;
				String uName       = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 8);
				String pswd        = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 9);
				String RANo        = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 10);
				String voidRemarks = rwe.getCellData("PAYLESS_VOIDRENTAL", k, 11);
				
				/* Open GUI URL's */
				functions.openURL(thinClient);
				
				/* Login */
				functions.login(uName, pswd);
				functions.navigateToTab("DisplayRental");
				Thread.sleep(1000);
				functions.displayRA(RANo);
				Thread.sleep(10000);
				
				try
				{
					Thread.sleep(5000);
				    String alertMsg = driver.switchTo().alert().getText();
				    driver.switchTo().alert().accept();
				    Thread.sleep(3000);
					rwe.setCellData("PAYLESS_VOIDRENTAL", k, 16, alertMsg);
					driver.findElement(By.id("rentalsearch:rentalSearchPopup:rentalSearchCancel")).click();
				}
			    catch(Exception e)
				{	
					WebElement RAStatusWebElement = driver.findElement(By.xpath("//*[@id='quickViewPanel']/div[2]/table/tbody/tr[1]/td/div[1]/div[2]/span"));
					WebDriverWait wait1 = new WebDriverWait(driver, 20);
					wait1.until(ExpectedConditions.visibilityOf(RAStatusWebElement));
					String RAStatus = RAStatusWebElement.getText();
					test = extent.createTest(testcasename);
					//String RAStatus = driver.findElement(By.cssSelector("span.qvTextDisplayStatus.qvDiStatus")).getText(); 
					if(RAStatus.equalsIgnoreCase("Open Rental"))
					{
						functions.makeRAVoid(voidRemarks);	
						Thread.sleep(10000);
						String db_VoidSuccessDialog = driver.findElement(By.xpath("//*[@id='successVoidDialog']/div/div/div[2]/div[1]/div[2]/div")).getText();
						String db_VoidRenterName    = driver.findElement(By.id("displayRentalMenuBarDialogs:successVoidForm:voidRenterName")).getText();
						String db_VoidRA            = driver.findElement(By.id("displayRentalMenuBarDialogs:successVoidForm:successVoidRA")).getText();
						String db_VoidSuccessOutMsg = driver.findElement(By.id("displayRentalMenuBarDialogs:successVoidForm:voidSuccessOutMsg")).getText();
						rwe.setCellData("PAYLESS_VOIDRENTAL", k, 12, db_VoidSuccessDialog);
						rwe.setCellData("PAYLESS_VOIDRENTAL", k, 13, db_VoidRenterName);
						rwe.setCellData("PAYLESS_VOIDRENTAL", k, 14, db_VoidRA);
						rwe.setCellData("PAYLESS_VOIDRENTAL", k, 15, db_VoidSuccessOutMsg);
						test.log(Status.PASS, "Pass");
						//driver.findElement(By.id("displayRentalMenuBarDialogs:successVoidForm:voidPopupViewRentalButton")).click();
						driver.findElement(By.name("displayRentalMenuBarDialogs:successVoidForm:voidPopupViewRentalButton")).click();
						}
					else
					{
						rwe.setCellData("PAYLESS_VOIDRENTAL", k, 16, RAStatus);
						test.log(Status.FAIL, "Fail");
					}
				}
				Thread.sleep(10000);
				functions.logout();
				functions.closeWindows();
			}	//end of conditional If
		}     //end of iteration For
		} finally {
			// TODO: handle finally clause
		extent.flush();
	}
		
	}  // end of main class
}
