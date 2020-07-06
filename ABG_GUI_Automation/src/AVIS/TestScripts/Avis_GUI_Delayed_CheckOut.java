package AVIS.TestScripts;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import pomClasses.GUIFunctions;
import pomClasses.ReadWriteExcel;

public class Avis_GUI_Delayed_CheckOut {

		public static void main(String[] args) throws Exception {
			// TODO Auto-generated method stub
			for (int k = 1; k<= 8; k++) 
			{
				AVIS.CommonFunctions.ReadWriteExcel rw = new AVIS.CommonFunctions.ReadWriteExcel("C:\\Downloads\\Selenium\\AVIS\\TestData\\AVIS_GUIDelayed_CheckOut.xlsx");
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
			//String TestUrlEnd = TestUrl+TestStn;
			driver.get("https://uat.ccrgservices.com/wizardgui/ui/wizard.jsf?mnemonic=jfk");
			driver.manage().window().maximize();
			AVIS.CommonFunctions.GUIFunctions functions = new AVIS.CommonFunctions.GUIFunctions(driver);
			//functions.link(tokenURL, thinClient);
			/* Login */
			functions.login(Testqauser, Testqapwd);
			//functions.login("qa.user", "Avis2018#");
			driver.navigate().refresh();
			
		
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[@id='searchString']")).sendKeys("0988-6083-US-4");
			Thread.sleep(5000);
			driver.findElement(By.xpath("//input[@id='searchCommandLink']")).click();
			Thread.sleep(5000);
			driver.findElement(By.xpath("//span[@id='delayBtn']")).click();
			//Ra number
			//
			//701215944
			driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:raNo']")).sendKeys("701215933");
			//agent Id
			driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:agentId']")).sendKeys("12345");
			//outdate
			//input[@id='menulist:checkoutContainer:checkoutForm:outdate1_hid']
			//checkoutdate
			driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:outdate1_hid']")).sendKeys("10/07/18");
			//checkintime
			driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:outdate2']")).sendKeys("10:00 AM");
			//MVA number
			driver.findElement(By.xpath("//div[@ng-show='!isOffline']//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys("");
			//mileage
			driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys("");
			
			//div[@id='ajaxStatusPanel']//div[@class='modal-body']
			//click on delay continue
			driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).click();
			driver.findElement(By.xpath("//div[@ng-show='!isOffline']//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys("64459010");
	        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
	        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
	        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys("1200");
	         
			}
		  }
		}
}
