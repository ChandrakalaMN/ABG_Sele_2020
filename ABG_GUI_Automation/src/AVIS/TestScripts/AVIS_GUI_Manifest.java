package AVIS.TestScripts;

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
import com.gui.report.Extentmanager;

import AVIS.CommonFunctions.*;
import AVIS.CommonFunctions.ReadWriteExcel;

/**
'#################################################################################################################################
'## SCRIPT NAME:                           Avis_GUI_Manifest
'## BRAND:                                         AVIS
'## DESCRIPTION:                           Perform a Manifest of different Types to capture the report details in different regions.
'## FUNCTIONAL AREA :                      Manifest Screen
'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
'## OUTPUT:                                Reservation should be created successfully.

'##################################################################################################################################
 **/

public class AVIS_GUI_Manifest
{

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
			
		   for (int k=1; k<=2; k++)
		   {
			   ReadWriteExcel rwe = new ReadWriteExcel("C:\\Backup\\ABG SE Scripts\\TestData\\Manifest\\Avis_GUITestData_Manifest_Regression.xlsx");
			   System.setProperty("webdriver.chrome.driver", "C:\\Backup\\ABG SE Scripts\\chromedriver.exe");
			   WebDriver driver = new ChromeDriver();
			   GUIFunctions functions = new GUIFunctions(driver);
			   String ScreenShotPath = "C:\\Selenium\\Screenshots\\Avis_Manifest\\";
			   
			   String Execute 					= rwe.getCellData("Manifest_Avis", k,  2);
			   String TESTCASENAME 			= rwe.getCellData("Manifest_Avis", k,  4);
		       String THINCLIENTURL 			= rwe.getCellData("Manifest_Avis", k,  6);
			   String USERID 					= rwe.getCellData("Manifest_Avis", k,  7);
			   String PASSWORD 				= rwe.getCellData("Manifest_Avis", k,  8);
			   String OUTSTATION				= rwe.getCellData("Manifest_Avis", k,  9);
//			   String COUNTER_NUM				= rwe.getCellData("Manifest_Avis", k, 13);
			   
			if (Execute.equals("N"))	
				driver.close();
			
			if (Execute.equals("Y"))
			{
				driver.get(THINCLIENTURL + OUTSTATION);
				functions.login(USERID, PASSWORD);

				driver.navigate().refresh();
				Thread.sleep(10000);
				driver.manage().window().maximize();
					
				driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
				Thread.sleep(2000);
					
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
				Thread.sleep(2000);
					
				String OneDay_LOR = rwe.getCellData("Manifest_Avis", k, 11);
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(OneDay_LOR);
				Thread.sleep(2000);
							
				driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
				Thread.sleep(8000);
				
				String DetailType = rwe.getCellData("Manifest_Avis", k, 10);
				System.out.println(DetailType);
					
			    driver.findElement(By.xpath("//*[@id=\"manifest_submenu\"]/div/div[3]/button")).click();;
			    Thread.sleep(3000);
			    System.out.println(" Manifest Type is " + DetailType);
				driver.findElement(By.linkText(DetailType)).click();
				Thread.sleep(3000);
					
				WebElement NoRes = driver.findElement(By.xpath("//*[@id='manifestGrid']/div[1]/div[2]/div/span"));
				NoRes.getText();
				String NoResFound = NoRes.getText();
				System.out.println(NoResFound);
				String ResVal	=	"No reservations found.";
				String NoData	=	"09-DATA NOT AVAILABLE FOR DTE ENTERED";
	
	// Logout the application when the "No reservations found"
				if(NoResFound.equals(ResVal) || NoResFound.equals(NoData))
				{
					Date d= new Date();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(src,new File( ScreenShotPath +TESTCASENAME+ sdf.format(d)+"_1Day"+".pgn"));
					
					functions.logout();
					driver.close();
				}
				else
				{
					WebElement TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					String TotItems1= TotItems.getText();
					String str = TotItems1;
				    System.out.println(str);
				 
				    // remove the last character 
				    str = str.substring(9, str.length()-5);
				    rwe.setCellData("Manifest_Avis", k, 14, str);
					Thread.sleep(2000);
			        		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					String ResNo1= ResNo.getText();
						
	//		Thread.sleep(2000);
					rwe.setCellData("Manifest_Avis", k, 15, ResNo1);
					
					WebElement AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String AWD1= AWD.getText();
					System.out.println(AWD.getText());	
					rwe.setCellData("Manifest_Avis", k, 17, AWD1);
					
					WebElement Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					String Wizard1 = Wizard.getText();
					rwe.setCellData("Manifest_Avis", k, 18, Wizard1);
					
					WebElement Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					String Coupon1 = Coupon.getText();
					rwe.setCellData("Manifest_Avis", k, 46, Coupon1);
					
					WebElement PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					String PickUpLoc1 = PickUpLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 19, PickUpLoc1);
					
					WebElement PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					PickupDate.getText();
					String PickupDate1 = PickupDate.getText();
					System.out.println(PickupDate.getText());	
					rwe.setCellData("Manifest_Avis", k, 20, PickupDate1);
					
					WebElement ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					ReturnLoc.getText();
					String ReturnLoc1 = ReturnLoc.getText();
					System.out.println(ReturnLoc.getText());	
					rwe.setCellData("Manifest_Avis", k, 21, ReturnLoc1);
					
					WebElement ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					ReturnDateTime.getText();
					String ReturnDateTime1 = ReturnDateTime.getText();
					System.out.println(ReturnDateTime.getText());	
					rwe.setCellData("Manifest_Avis", k, 22, ReturnDateTime1);
					
					WebElement CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					CarClass.getText();
					String CarClass1 = CarClass.getText();
					System.out.println(CarClass.getText());	
					rwe.setCellData("Manifest_Avis", k, 23, CarClass1);
					
					WebElement BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					BookingDt.getText();
					String BookingDt1 = BookingDt.getText();
					System.out.println(BookingDt.getText());	
					rwe.setCellData("Manifest_Avis", k, 24, BookingDt1);
					Thread.sleep(2000);
					
	// Capture the screen shot of the report
					Date d1= new Date();
					SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src1= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					
					FileUtils.copyFile(src1,new File( ScreenShotPath +TESTCASENAME+ sdf1.format(d1)+"_1Day"+".pgn"));
					
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
	
	// Max days (2 days)of Manifest
					driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
					Thread.sleep(2000);
					
					String TwoDay_LOR = rwe.getCellData("Manifest_Avis", k, 12);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(TwoDay_LOR);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
					Thread.sleep(15000);
	
	//Capture the total number of items are displayed		
					WebElement max_TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					String max_TotItems1= max_TotItems.getText();
										
					String max_str = max_TotItems1;
					max_str = max_str.substring(9, max_str.length()-5);
					rwe.setCellData("Manifest_Avis", k, 25, max_str);
	
	//Select a one of the record to open in report format		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement max_ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					String max_ResNo1= max_ResNo.getText();
					rwe.setCellData("Manifest_Avis", k, 26, max_ResNo1);
					
					WebElement max_AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String max_AWD1= max_AWD.getText();
					rwe.setCellData("Manifest_Avis", k, 28, max_AWD1);
					
					WebElement max_Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					String max_Wizard1 = max_Wizard.getText();
					rwe.setCellData("Manifest_Avis", k, 29, max_Wizard1);
					
					WebElement max_Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					max_Coupon.getText();
					
					WebElement max_PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					String max_PickUpLoc1 = max_PickUpLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 30, max_PickUpLoc1);
					
					WebElement max_PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					String max_PickupDate1 = max_PickupDate.getText();
					rwe.setCellData("Manifest_Avis", k, 31, max_PickupDate1);
					
					WebElement max_ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					String max_ReturnLoc1 = max_ReturnLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 32, max_ReturnLoc1);
					
					WebElement max_ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					String max_ReturnDateTime1 = max_ReturnDateTime.getText();
					rwe.setCellData("Manifest_Avis", k, 33, max_ReturnDateTime1);
					
					WebElement max_CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					String max_CarClass1 = max_CarClass.getText();
					rwe.setCellData("Manifest_Avis", k, 34, max_CarClass1);
					
					WebElement max_BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					String max_BookingDt1 = max_BookingDt.getText();
					rwe.setCellData("Manifest_Avis", k, 35, max_BookingDt1);
					Thread.sleep(2000);
					
	// Capture the screen shot of the report
					Date d2= new Date();
					SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_2Day"+".pgn"));
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
					
	// Logout the application
					test = extent.createTest(TESTCASENAME);
					
					if (rwe.getCellData("Manifest_Avis", k, 22).isEmpty())
					{
						rwe.setCellData("Manifest_Avis", k, 23, "FAIL");
						test.log(Status.FAIL, "Fail");
					}
					else
					{
						test.log(Status.PASS, "Pass");
						rwe.setCellData("Manifest_Avis", k, 23, "PASS");
					}

					
					functions.logout();
					driver.close();
				}
			}
			
			else
			{
			   System.out.println("Execution status is N for iteration "+k+"...");
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