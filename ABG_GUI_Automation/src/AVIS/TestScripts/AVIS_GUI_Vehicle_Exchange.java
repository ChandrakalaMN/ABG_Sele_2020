/**
 * '#############################################################################################################################
 * SCRIPT NAME      :  AVIS_GUI_Vehicle_Exchange 
 * BRAND            :  AVIS
 * DESCRIPTION      :  Vehicle Exchange - On the lot and off the lot exchange 
 * FUNCTIONAL AREA  :  Display Rental Screen 
 * PRECONDITION     :  All the required Test Data should be available in Test Data Sheet. 
 * OUTPUT           :  Vehicle Exchange transaction should be completed successfully.
 * '#############################################################################################################################
 **/

package AVIS.TestScripts;

import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;

public class AVIS_GUI_Vehicle_Exchange
{
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
		try 
		{
		// Read input from excel
		for (int k = 1; k <= 30; k++)
		{
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Selenium\\TestData\\GUI_AVIS_Vehicle_Exchange.xlsx");
			String Execute = rwe.getCellData("INPUT_DATA", k, 2);
			if (Execute.equals("Y"))
			{
				System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver.exe");
				ChromeDriver driver = new ChromeDriver();
				GUIFunctions functions = new GUIFunctions(driver);
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				System.out.println(" Test Case No " + k);
				String testCaseName= rwe.getCellData("INPUT_DATA", k, 5);
				String clientURL   = rwe.getCellData("INPUT_DATA", k, 6);
				String outSTA      = rwe.getCellData("INPUT_DATA", k, 9);
				String thinClient  = clientURL+outSTA;
				String uName       = rwe.getCellData("INPUT_DATA", k, 7);
				String pswd        = rwe.getCellData("INPUT_DATA", k, 8);
				String raNumber    = rwe.getCellData("INPUT_DATA", k, 10);
				String exchngOpt   = rwe.getCellData("INPUT_DATA", k, 11);
				String MVAnumber   = rwe.getCellData("INPUT_DATA", k, 12);
				String Mileage     = rwe.getCellData("INPUT_DATA", k, 13);
				String Fuel        = rwe.getCellData("INPUT_DATA", k, 14);
//				String delayed     = rwe.getCellData("INPUT_DATA", k, 15);
				String mileageIn   = rwe.getCellData("INPUT_DATA", k, 16);
				String prchsfuel   = rwe.getCellData("INPUT_DATA", k, 17);
				String fuelin      = rwe.getCellData("INPUT_DATA", k, 18);
				String adjstmnt    = rwe.getCellData("INPUT_DATA", k, 19);
				String vehdamaged  = rwe.getCellData("INPUT_DATA", k, 20);
				String accreport   = rwe.getCellData("INPUT_DATA", k, 21);
				String fuelother   = rwe.getCellData("INPUT_DATA", k, 22);
				String reason      = rwe.getCellData("INPUT_DATA", k, 23);
				String amount      = rwe.getCellData("INPUT_DATA", k, 24);
				String oneway      = rwe.getCellData("INPUT_DATA", k, 25);
				String misc        = rwe.getCellData("INPUT_DATA", k, 26);
				String miscfee     = rwe.getCellData("INPUT_DATA", k, 27);
				String VEremarks   = rwe.getCellData("INPUT_DATA", k, 28);	
				String dlydInDate  = rwe.getCellData("INPUT_DATA", k, 29);	
				String dlydInTime  = rwe.getCellData("INPUT_DATA", k, 30);	
				String methpay     = rwe.getCellData("INPUT_DATA", k, 31);	
				
				functions.link(thinClient);

				/* Login */
				functions.login(uName, pswd);
				functions.navigateToTab("DisplayRental");
				Thread.sleep(500);
				functions.enterSearchRAResMVAnumber(raNumber);
				functions.clickSearchBtnDisplayRAResMVA();
				Thread.sleep(10000);
				functions.exchangeOptions();
				functions.onoffthelot(exchngOpt);
				
				/* Logic for on the lot and off the lot exchange */
				if (rwe.getCellData("INPUT_DATA", k, 11).equalsIgnoreCase("On the Lot Exchange"))
				{
					functions.onthelottransaction(MVAnumber, Mileage, Fuel);
					functions.onthelotclickokBtn(methpay);
				} 
				else if (rwe.getCellData("INPUT_DATA", k, 11).equalsIgnoreCase("Off the Lot Exchange"))
				{
					if (rwe.getCellData("INPUT_DATA", k, 15).equalsIgnoreCase("Yes"))
					{
						functions.delayedoffthelot(dlydInDate, dlydInTime, mileageIn, prchsfuel, fuelin, adjstmnt, vehdamaged, accreport, fuelother, reason, amount, oneway, misc, miscfee, MVAnumber, Mileage, Fuel, VEremarks);
					
					} else if (rwe.getCellData("INPUT_DATA", k, 15).equalsIgnoreCase("No"))
					{
						functions.notdelayedoffthelot(mileageIn, prchsfuel, fuelin, adjstmnt, vehdamaged, accreport, fuelother, reason, amount, oneway, misc, miscfee, MVAnumber, Mileage, Fuel, VEremarks, methpay);
					}
				}
				
				Thread.sleep(10000);
				
				String Finalmsg = driver.findElement(By.id("modifyRentalDialog:successModDialogForm:successName")).getText();
				rwe.setCellData("INPUT_DATA", k, 32, Finalmsg);
				
				String FinalmsgRA = driver.findElement(By.id("modifyRentalDialog:successModDialogForm:successRA")).getText();
				rwe.setCellData("INPUT_DATA", k, 33, FinalmsgRA);
				
				String FinalSuccessMsg = driver.findElement(By.id("modifyRentalDialog:successModDialogForm:successOutMsg")).getText();
				rwe.setCellData("INPUT_DATA", k, 34, FinalSuccessMsg);
				
				driver.findElement(By.id("modifyRentalDialog:successModDialogForm:rentalModCompDone")).click();
				
				Thread.sleep(5000);
				functions.logout();
				Thread.sleep(1000);
				functions.closeWindows();
				
				test = extent.createTest(testCaseName);

				if (rwe.getCellData("INPUT_DATA", k, 32).isEmpty()) 
				{
					rwe.setCellData("INPUT_DATA", k, 36, "FAIL");
					test.log(Status.FAIL, "Fail");
				} else 
				{
					test.log(Status.PASS, "Pass");
					rwe.setCellData("INPUT_DATA", k, 36, "PASS");
				}
			}
		}
	} finally 
		{
		// TODO: handle finally clause
		extent.flush();
	    }
	}
}