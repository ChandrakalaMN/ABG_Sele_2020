package Payless.TestScripts;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import com.gui.report.Extentmanager;
import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;

/**
'#############################################################################################################################
'## SCRIPT NAME:                           GUI_Manifest_Payless
'## BRAND:                                         Payless
'## DESCRIPTION:                           Create Reservation for daily LOR with different insurance and counter products.
'## FUNCTIONAL AREA :                      Reservation Rates Screen
'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
'## OUTPUT:                                Reservation should be created successfully.


HISTORY 
05-SEP-2018 - GUIFunctions class created for GUI Common functionalities and CR functionality 
'#############################################################################################################################
 **/

public class PAYLESS_GUI_Manifest
{
	public static void main(String[] args) throws Throwable {
		
		for (int k=1; k<=6; k++)
		{
			ReadWriteExcel rwe = new ReadWriteExcel("C:\\Backup\\ABG SE Scripts\\TestData\\Manifest\\PAYLESS_GUITestData_Manifest_Regression.xlsx");
			System.setProperty("webdriver.chrome.driver", "C:\\Backup\\ABG SE Scripts\\chromedriver.exe");
			WebDriver driver = new ChromeDriver();
				
				String Execute 					= rwe.getCellData("INPUT_DATA", k, 	2);
				String THINCLIENTURL 			= rwe.getCellData("INPUT_DATA", k, 	7);
				String USERID 					= rwe.getCellData("INPUT_DATA", k, 	8);
				String PASSWORD 				= rwe.getCellData("INPUT_DATA", k, 	9);
				String OUTSTATION				= rwe.getCellData("INPUT_DATA", k, 10);
//				String COUNTER_NUM				= rwe.getCellData("INPUT_DATA", k, 14);
			
			if ( Execute.equals("N") )	
				driver.close();
			
			if ( Execute.equals("Y") )
			{
				driver.get(THINCLIENTURL + OUTSTATION);
				String tokenURL = rwe.getCellData("INPUT_DATA", k, 6);
				System.out.println(tokenURL);			
				driver.findElement(By.name("username")).sendKeys(USERID);
				driver.findElement(By.name("PASSWORD")).sendKeys(PASSWORD);
					
				driver.findElement(By.cssSelector("body > table:nth-child(7) > tbody > tr > td:nth-child(2) > form > table > tbody > tr:nth-child(3) > td > input[type=\"SUBMIT\"]:nth-child(3)")).click();
				driver.navigate().refresh();
				Thread.sleep(10000);
					
				driver.manage().window().maximize();
					
				driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
				Thread.sleep(2000);
					
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
				Thread.sleep(2000);
					
				String OneDay_LOR = rwe.getCellData("INPUT_DATA", k, 12);
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(OneDay_LOR);
				Thread.sleep(2000);
							
				driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
				Thread.sleep(8000);
				
				String DetailType = rwe.getCellData("INPUT_DATA", k, 11);
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
				
				if(NoResFound.equals(ResVal) || NoResFound.equals(NoData)){
					driver.findElement(By.id("headerLogOutButton")).click();
					Thread.sleep(2000);
					driver.findElement(By.id("logoutForm:yesLogoutButton")).click();
					Thread.sleep(3000);
					driver.close();
				}else{
		    		
					WebElement TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					TotItems.getText();
					String TotItems1= TotItems.getText();
					String str = TotItems1;
				    System.out.println(str);
				 
				    // remove the last character 
				    str = str.substring(9, str.length()-5);
				    rwe.setCellData("INPUT_DATA", k, 15, str);
					Thread.sleep(2000);
			        		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					ResNo.getText();
					String ResNo1= ResNo.getText();
					System.out.println(ResNo.getText());	
	//		Thread.sleep(2000);
					rwe.setCellData("INPUT_DATA", k, 16, ResNo1);
					
					WebElement AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String AWD1= AWD.getText();
					System.out.println(AWD.getText());	
					rwe.setCellData("INPUT_DATA", k, 18, AWD1);
					
					WebElement Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					Wizard.getText();
					String Wizard1 = Wizard.getText();
					System.out.println(Wizard.getText());	
					rwe.setCellData("INPUT_DATA", k, 19, Wizard1);
					
					WebElement Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					Coupon.getText();
					String Coupon1 = Coupon.getText();
					System.out.println(Coupon.getText());	
					rwe.setCellData("INPUT_DATA", k, 47, Coupon1);
					
					WebElement PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					PickUpLoc.getText();
					String PickUpLoc1 = PickUpLoc.getText();
					System.out.println(PickUpLoc.getText());	
					rwe.setCellData("INPUT_DATA", k, 20, PickUpLoc1);
					
					WebElement PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					PickupDate.getText();
					String PickupDate1 = PickupDate.getText();
					System.out.println(PickupDate.getText());	
					rwe.setCellData("INPUT_DATA", k, 21, PickupDate1);
					
					WebElement ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					ReturnLoc.getText();
					String ReturnLoc1 = ReturnLoc.getText();
					System.out.println(ReturnLoc.getText());	
					rwe.setCellData("INPUT_DATA", k, 22, ReturnLoc1);
					
					WebElement ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					ReturnDateTime.getText();
					String ReturnDateTime1 = ReturnDateTime.getText();
					System.out.println(ReturnDateTime.getText());	
					rwe.setCellData("INPUT_DATA", k, 23, ReturnDateTime1);
					
					WebElement CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					CarClass.getText();
					String CarClass1 = CarClass.getText();
					System.out.println(CarClass.getText());	
					rwe.setCellData("INPUT_DATA", k, 24, CarClass1);
					
					WebElement BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					BookingDt.getText();
					String BookingDt1 = BookingDt.getText();
					System.out.println(BookingDt.getText());	
					rwe.setCellData("INPUT_DATA", k, 25, BookingDt1);
					Thread.sleep(2000);
					
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
	
	// Max days (2 days)of Manifest
					driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
					Thread.sleep(2000);
					
					String TwoDay_LOR = rwe.getCellData("INPUT_DATA", k, 13);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(TwoDay_LOR);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
					Thread.sleep(15000);
	
	//Capture the total number of items are displayed		
					WebElement max_TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					max_TotItems.getText();
					System.out.println(max_TotItems.getText());
					String max_TotItems1= max_TotItems.getText();
					System.out.println(max_TotItems.getText());
					
					String max_str = max_TotItems1;
					System.out.println(max_str);
				    max_str = max_str.substring(9, max_str.length()-5);
					rwe.setCellData("INPUT_DATA", k, 26, max_str);
	
	//Select a one of the record to open in report format		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement max_ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					max_ResNo.getText();
					String max_ResNo1= max_ResNo.getText();
					System.out.println(max_ResNo.getText());	
					rwe.setCellData("INPUT_DATA", k, 27, max_ResNo1);
					
					WebElement max_AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String max_AWD1= max_AWD.getText();
					System.out.println(max_AWD.getText());	
					rwe.setCellData("INPUT_DATA", k, 29, max_AWD1);
					
					WebElement max_Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					max_Wizard.getText();
					String max_Wizard1 = max_Wizard.getText();
					System.out.println(max_Wizard.getText());	
					rwe.setCellData("INPUT_DATA", k, 30, max_Wizard1);
					
					WebElement max_Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					max_Coupon.getText();
					
					WebElement max_PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					max_PickUpLoc.getText();
					String max_PickUpLoc1 = max_PickUpLoc.getText();
					System.out.println(max_PickUpLoc.getText());	
					rwe.setCellData("INPUT_DATA", k, 31, max_PickUpLoc1);
					
					WebElement max_PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					max_PickupDate.getText();
					String max_PickupDate1 = max_PickupDate.getText();
					System.out.println(max_PickupDate.getText());	
					rwe.setCellData("INPUT_DATA", k, 32, max_PickupDate1);
					
					WebElement max_ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					ReturnLoc.getText();
					String max_ReturnLoc1 = max_ReturnLoc.getText();
					System.out.println(max_ReturnLoc.getText());	
					rwe.setCellData("INPUT_DATA", k, 33, max_ReturnLoc1);
					
					WebElement max_ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					max_ReturnDateTime.getText();
					String max_ReturnDateTime1 = max_ReturnDateTime.getText();
					System.out.println(max_ReturnDateTime.getText());	
					rwe.setCellData("INPUT_DATA", k, 34, max_ReturnDateTime1);
					
					WebElement max_CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					max_CarClass.getText();
					String max_CarClass1 = max_CarClass.getText();
					System.out.println(max_CarClass.getText());	
					rwe.setCellData("INPUT_DATA", k, 35, max_CarClass1);
					
					WebElement max_BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					max_BookingDt.getText();
					String max_BookingDt1 = max_BookingDt.getText();
					System.out.println(max_BookingDt.getText());
					rwe.setCellData("INPUT_DATA", k, 36, max_BookingDt1);
					Thread.sleep(2000);
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
			
	// Logout the application		
					driver.findElement(By.xpath("//*[@id=\"headerLogOutButton\"]")).click();
					Thread.sleep(2000);
					
					driver.findElement(By.xpath("//*[@id=\"logoutForm:yesLogoutButton\"]")).click();
					Thread.sleep(3000);
					
					driver.close();
				}
			}else{
			System.out.println("Execution status is N for iteration "+k+"...");
		}
	}
}}