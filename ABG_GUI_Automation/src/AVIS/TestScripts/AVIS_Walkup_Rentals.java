package AVIS.TestScripts;


import java.io.File;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.formula.functions.Function;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.gui.report.Extentmanager;

import AVIS.CommonFunctions.*;
import AVIS.CommonFunctions.ReadWriteExcel;

public class AVIS_Walkup_Rentals
{
	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport()
	{
		extent = Extentmanager.GetExtent();
		//test = extent.createTest("GUI");
	}

//	public static void main(String[] args) throws Exception {
	@Test
	public void test() throws Exception {
		// TODO Auto-generated method stub
		try {
			
		for (int k = 1; k<= 8; k++) 
		{
			ReadWriteExcel rw = new ReadWriteExcel("C:\\Downloads\\Selenium\\AVIS\\TestData\\AVIS_GUITestData_WalkUp.xlsx");
			String TestExecute = rw.getCellData("Avis_GUI", k ,2);
		
			if (TestExecute.equals("Y"))
			{	
			String Testcasename   =rw.getCellData("Avis_GUI", k, 4);
			String TestUrl        = rw.getCellData("Avis_GUI", k, 7);
			String TestStn        = rw.getCellData("Avis_GUI", k, 8);
			String Testqauser     = rw.getCellData("Avis_GUI", k, 9);
			String Testqapwd      = rw.getCellData("Avis_GUI", k, 10);
			String LastName       = rw.getCellData("Avis_GUI", k, 11);
			String FirstName      = rw.getCellData("Avis_GUI", k, 12);
			String CheckOutDate   = rw.getCellData("Avis_GUI", k, 13);
			String CheckOutTime   = rw.getCellData("Avis_GUI", k, 14);
			String InStation      = rw.getCellData("Avis_GUI", k, 15);
			String CheckInDate    = rw.getCellData("Avis_GUI", k, 16);
			String CheckInTime    = rw.getCellData("Avis_GUI", k, 17);
			String CarGroup       = rw.getCellData("Avis_GUI", k, 18);
			String Awd            = rw.getCellData("Avis_GUI", k, 19);
			String CreditDeb_Card = rw.getCellData("Avis_GUI", k, 20);
			String CardNumber     = rw.getCellData("Avis_GUI", k, 21);
			String Month          = rw.getCellData("Avis_GUI", k, 22);
			String Year           = rw.getCellData("Avis_GUI", k, 23);
			String Reason         = rw.getCellData("Avis_GUI", k, 24);
			String Insurance      = rw.getCellData("Avis_GUI", k, 25);
			String Counterproduct = rw.getCellData("Avis_GUI", k, 26);

		//System.out.println("Test case name is :" +TestCaseName);
		Thread.sleep(5000);
		System.setProperty("webdriver.chrome.driver", "C:\\Downloads\\Selenium\\chromedriver.exe");
		ChromeOptions options = new ChromeOptions(); 
		options.addArguments("disable-infobars"); 
		WebDriver driver=new ChromeDriver(options);
		String TestUrlEnd = TestUrl+TestStn;
		driver.get(TestUrlEnd);
		GUIFunctions functions = new GUIFunctions(driver);
		//functions.link(tokenURL, thinClient);
		/* Login */
		functions.login(Testqauser, Testqapwd);
		driver.navigate().refresh();
		driver.manage().window().maximize();
	
		Thread.sleep(5000);
		//functions.link(tokenURL, thinClient);
		/* Login */
		//functions.login("qa.user", "Avis2018#");
		//.setCellData("Walkup_Rental", 1,29,TestCaseName);
		driver.findElement(By.xpath("//input[@id='rateShopCommandButton']")).click();
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:pickupDate_hid']")).sendKeys(CheckOutDate);
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:pickupTime']")).sendKeys(CheckOutTime);
		Thread.sleep(7000);
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:returnDate_hid']")).sendKeys(CheckInDate);
		Thread.sleep(1000);
		
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:returnTime']")).sendKeys(CheckInTime);
		Thread.sleep(1000);
		driver.findElement(By.xpath("//div[@class='form-horizontal']//div[@class='col-xs-12 col-sm-12 col-md-12 col-lg-12']")).click();
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:wizardNumber']")).sendKeys("");
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:coupon']")).sendKeys("");
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:qty']")).sendKeys("");
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:rateShop']")).click();
		Thread.sleep(9000);
		
		String Car_val =driver.findElement(By.xpath("//*[@id='resRateLookupDlg:rateLookupForm:groups_data']")).getAttribute("innerText"); 
		String[] temp_Car_Val = Car_val.split("\\n");
	     int len_tempCar= temp_Car_Val.length;
	     String Excel_ValCar = CarGroup; 
	     for (int i=1;i< len_tempCar;i++){
	         // System.out.println(temp_Car_Val[i]);
	          String[] Ary_Cls_name = temp_Car_Val[i].split("-");
	          String Cls_name = Ary_Cls_name[0].trim();
	        //  System.out.println(Cls_name);
	          if (Excel_ValCar.equals(Cls_name) ){
	              System.out.println(i);
	              int j=i-1;
	              String Rate = "//input[@value='standard,"+j+"']";
	            		//input[@value='standard,0']
	              driver.findElement(By.xpath(Rate)).click();
	            		  break;
	              
	          }
	                
	     }
		
		
		driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:checkOutButton']")).click();
		// walk up
		//driver.findElement(By.xpath("//input[@id='walkupCommandwalkup']")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkuplastname']")).sendKeys(LastName);
		driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkupfirstname']")).sendKeys(FirstName);
		// if wizard number needs to be given as input  
		//driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkupwizardnumber']")).sendKeys("");
		driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkupContinue']")).click();
		Thread.sleep(3000);
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys("US");
		 Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys("NY");
		 Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys("123456");
		 Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys("08/23/89");
		 Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys("abg");
		 Thread.sleep(2000);
		//Address 1
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys("12,Ermro");
		//Address 2
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys("5th coeswe way ");
		//Address 3
		 Thread.sleep(7000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys("Parsippany,NJ,07054,US");
		//Contact Info
		 Thread.sleep(7000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys("9924566446");
		//E-Mail 
		 Thread.sleep(7000);
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys("abgs@abg.com");
		//Freq.Travel
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).sendKeys("");
		//Number
		driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftNumber']")).sendKeys("");
		JavascriptExecutor js = (JavascriptExecutor) driver;  
		   js.executeScript("window.scrollBy(0,400)");
		   Thread.sleep(2000);
		
		   WebElement CheckOut_lstCOCCDC = driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType"));
           Select MOPDropDownList    = new Select(CheckOut_lstCOCCDC);
           MOPDropDownList.selectByVisibleText(CreditDeb_Card );
           Thread.sleep(3000);
           driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber")).sendKeys(CardNumber); 
           Thread.sleep(3000);            
           driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth")).sendKeys(Month); 
           Thread.sleep(3000);
           driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear")).sendKeys(Year); 
           Thread.sleep(3000);            
           WebElement CheckOut_MOPReason = driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason"));
           Select MOPReasonDropDownList  = new Select(CheckOut_MOPReason);
           MOPReasonDropDownList.selectByVisibleText(Reason);
           Thread.sleep(3000);
           
           // Vehicle Selection
           driver.findElement(By.xpath("//div[@ng-show='!isOffline']//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys("64459010");
           driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys("1200");
           driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:checkoutSpace']")).sendKeys("");
           driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).click();
           Thread.sleep(10000);
           
           //driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("0");
           driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
           //checkoutRepromptDialog:repromptForm:repromptSubmitButton
           Thread.sleep(10000);
           //driver.findElement(By.id("id=\"footerForm:completeCheckoutButton\"")).click();
           driver.findElement(By.id("footerForm:completeCheckoutButton")).click();
           Thread.sleep(10000);
           driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();               
           Thread.sleep(7000); 

           if(CreditDeb_Card.equals("American Express"))
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
           
           Thread.sleep(10000);
           //driver.findElement(By.id("id=printPreviewDialog:printPreviewForm:closePrintPreviewBtn")).click();
           Thread.sleep(10000);
           
           
           
           
		   
           String strCOLNFNGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 29, strCOLNFNGetText);
           System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
           String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 28, strCOVehicleModelGetText);
           System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
           String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 27, strCOResNoGetText);
           System.out.println(" Reservation No value is " + strCOResNoGetText);
           String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 30, strCOResNoGetText);
           System.out.println(" MVA No value is " + strCOMVAGetText);
           String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 31, strCORANoGetText);
           System.out.println(" RA Number value is " + strCORANoGetText);
           String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 32, strCOLicensePlateNoGetText);
           System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
           String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
           Thread.sleep(7000);
           rw.setCellData("Avis_GUI", k, 33, strCOEstimatedCostGetText);
           System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
           
           String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
           Thread.sleep(7000);
           System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
           
	        String ScreenShotPath = "C:\\Downloads\\Selenium\\Screenshots\\";
   			//String testcasename ="TC1_walkup_dd";
   			File file=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	   		FileUtils.copyFile(file,new File( ScreenShotPath +Testcasename+ ".pgn"));
           
           driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
           Thread.sleep(7000);
		
           //sdf.format(d)+
           driver.close();
        
		}	
		else
		{
			System.out.println(k);
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
