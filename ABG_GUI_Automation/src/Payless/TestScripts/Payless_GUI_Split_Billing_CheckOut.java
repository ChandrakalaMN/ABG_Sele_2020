/**
'#############################################################################################################################
'## SCRIPT NAME     : Payless_GUI_Split_Billing_CheckOut
'## BRAND           : Payless
'## DESCRIPTION     : Split Billing Checkout.
'## FUNCTIONAL AREA : Checkout
'## PRECONDITION    : All the required Test Data should be available in Test Data Sheet.
'## OUTPUT          : Split billing checkout should be successful.
'#############################################################################################################################
 **/

package Payless.TestScripts;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
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

public class Payless_GUI_Split_Billing_CheckOut
{
   ExtentReports extent; //TestNG Extent Reports Changes.
   ExtentTest test;     //TestNG Extent Reports Changes.
   
   @BeforeTest //TestNG Extent Reports Changes.
   public void startReport() //TestNG Extent Reports Changes.
   {
		//extent = Extentmanager.GetExtent(this.getClass().getSimpleName()); //TestNG Extent Reports Changes.
		//test = extent.createTest("GUI");
		extent = Extentmanager.GetExtent();
   }
   
   //public static void main(String[] args) throws IOException, Exception, FileNotFoundException
   
	@Test //TestNG Extent Reports Changes.
	public void test() throws Exception //TestNG Extent Reports Changes.
	{
	   try //TestNG Extent Reports Changes.
	   {
		  // Read input from excel
	      // need to find the no of iterations / iterations count.
	      
	      ChromeDriver driver = null;
	      
	      ReadWriteExcel rwe  = new ReadWriteExcel("C:\\GitHub\\ABG\\Payless\\TestData\\InternetGUI\\Payless_GUI_Split_Billing_CheckOut.xlsx");
	      int intRowCount     = rwe.sheet.getLastRowNum();
	      
	      System.out.println(" Total row count in Split Billing test data sheet is : " + intRowCount);
	      
	      for (int k=1;k<=intRowCount;k++)
	      { 
             String Execute = rwe.getCellData("INPUT_DATA",k,2);
             
		     if (Execute.equals("Y"))
             {
			    String outSTA = rwe.getCellData("INPUT_DATA", k, 10);
			    
			    System.setProperty("webdriver.chrome.driver","C:\\Program Files (x86)\\Selenium\\WebDrivers\\ChromeDrivers\\chromedriver.exe");
			    
			    driver = new ChromeDriver();
			    driver.manage().window().maximize();
			    driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
     		    System.out.println(" Current iteration : " + k);
			    //String tokenURL = rwe.getCellData("INPUT_DATA", k, 6); //For HWY-ATR ThinClient.
     		    //driver.get(tokenURL);
     		    //Thread.sleep(2000);
     		    //String WindowHandle = driver.getWindowHandle();
     		    //Robot robot = new Robot();
     		    //robot.keyPress(KeyEvent.VK_CONTROL);
     		    //robot.keyPress(KeyEvent.VK_T);
     		    //robot.keyRelease(KeyEvent.VK_CONTROL);
     		    //robot.keyRelease(KeyEvent.VK_T);
     		    //ArrayList<String> tabs = new ArrayList(driver.getWindowHandles());
     		    //driver.switchTo().window(tabs.get(1));
     		    //Thread.sleep(500);
			    String clientURL = rwe.getCellData("INPUT_DATA", k, 7);
			    
			    try
			    {
			       String thinClient = clientURL+outSTA;
			       driver.get(thinClient);
			       //Assert.assertTrue(driver.getTitle().equals("Login")); // Check whether ThinClient page with title Login is loaded.
			       //Assert.assertTrue(driver.getTitle().equals("Access Manager Login")); // Check whether Internet GUI login page with title Access Manager Login is loaded.
			       Assert.assertTrue(driver.getTitle().contains("Login"));
			    }
			    catch(Exception e)
			    {
			       System.out.println(" Login page not found. ");
			    }
                
			    //Login to GUI using Internet GUI URL.
			    
			    String strUserName = rwe.getCellData("INPUT_DATA", k, 8);
			    String strPassword = rwe.getCellData("INPUT_DATA", k, 9);
			    driver.findElement(By.name("username")).sendKeys(strUserName);
			    driver.findElement(By.name("PASSWORD")).sendKeys(strPassword);
			    driver.findElement(By.xpath("//input[@value='Login']")).click();
			    Thread.sleep(10000);
              
			    String pageTitle = driver.getTitle();
			    
		        if(pageTitle.equals("Avis Budget Group"))
			    {
			       System.out.println(pageTitle);
			    }
		        else
			    {
			       driver.navigate().refresh();
		        }
                
		        //Reading values from Excel.
                
		        String strTestCaseName               = rwe.getCellData("INPUT_DATA", k, 4);
		      //String strRETURNSTATION              = rwe.getCellData("INPUT_DATA", k, 34);
		        String strRETURNDATE                 = rwe.getCellData("INPUT_DATA", k, 35);
		      //String strRETURNTIME                 = rwe.getCellData("INPUT_DATA", k, 36);
     	        String strLastName                   = rwe.getCellData("INPUT_DATA", k, 16);
     	        String strFirstName                  = rwe.getCellData("INPUT_DATA", k, 17);		       
		        String strDRLICCountry               = rwe.getCellData("INPUT_DATA", k, 18);
		        String strDRLICState                 = rwe.getCellData("INPUT_DATA", k, 19);
		        String strDRLICNo                    = rwe.getCellData("INPUT_DATA", k, 20);
		        String strDRDOB                      = rwe.getCellData("INPUT_DATA", k, 21);
		        String strDRCompany                  = rwe.getCellData("INPUT_DATA", k, 22);
		        String strDRAddressLine1             = rwe.getCellData("INPUT_DATA", k, 23);
		        String strDRAddressLine2             = rwe.getCellData("INPUT_DATA", k, 24);
		        String strDRAddressLine3             = rwe.getCellData("INPUT_DATA", k, 25);
		        String strDRContactInfo              = rwe.getCellData("INPUT_DATA", k, 26);
		        String strDREmailID                  = rwe.getCellData("INPUT_DATA", k, 27);
		      //String strCOMOP                      = rwe.getCellData("INPUT_DATA", k, 28);
		        String strCOCCDC                     = rwe.getCellData("INPUT_DATA", k, 29);
		        String strCARDNO                     = rwe.getCellData("INPUT_DATA", k, 30);
		        String strCOMM                       = rwe.getCellData("INPUT_DATA", k, 31);
		        String strCOYY                       = rwe.getCellData("INPUT_DATA", k, 32);
		        String strMOPREASON                  = rwe.getCellData("INPUT_DATA", k, 33);
		        String strCARGROUP                   = rwe.getCellData("INPUT_DATA", k, 38);
		      //String strREMARKS                    = rwe.getCellData("INPUT_DATA", k, 43);
		        String strCOMVA                      = rwe.getCellData("INPUT_DATA", k, 44);
		        String strCOMILEAGE                  = rwe.getCellData("INPUT_DATA", k, 45);
		        String strCO_SB_MOP                  = rwe.getCellData("INPUT_DATA", k, 47);
		        String strCO_SB_CCDC                 = rwe.getCellData("INPUT_DATA", k, 48);
		        String strCO_SB_CARD_NO              = rwe.getCellData("INPUT_DATA", k, 49);
		        String strCO_SB_CARD_EXP_MM          = rwe.getCellData("INPUT_DATA", k, 50);
		        String strCO_SB_CARD_EXP_YY          = rwe.getCellData("INPUT_DATA", k, 51);
		        String strCO_SB_MOP_REASON           = rwe.getCellData("INPUT_DATA", k, 52);
		        String strCO_SB_DIFF_ADDRESS_OPTION  = rwe.getCellData("INPUT_DATA", k, 53);
		        String strCO_SB_DIFF_ADDR1           = rwe.getCellData("INPUT_DATA", k, 54);
		        String strCO_SB_DIFF_ADDR2           = rwe.getCellData("INPUT_DATA", k, 55);
		        String strCO_SB_DIFF_ADDR3           = rwe.getCellData("INPUT_DATA", k, 56);
		        String strCO_SB_NO_OF_DAYS_OPTION    = rwe.getCellData("INPUT_DATA", k, 57);
		      //String strCO_SB_UPSELL_OPTION        = rwe.getCellData("INPUT_DATA", k, 58);
		      //String strCO_SB_TANDM_POINTS_OPTION  = rwe.getCellData("INPUT_DATA", k, 59);
		        String strCO_SB_DOLLAR_AMOUNT_OPTION = rwe.getCellData("INPUT_DATA", k, 60);
		        
		        driver.findElement(By.xpath("//input[@id='rateShopCommandButton']")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:returnDate_hid']")).sendKeys(strRETURNDATE);
		        Thread.sleep(1000);
		        String strRETURNTIME =  driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:pickupTime']")).getAttribute("value");
		        System.out.println(" return time " + strRETURNTIME);
		        driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:returnTime']")).sendKeys(strRETURNTIME);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//div[@class='form-horizontal']//div[@class='col-xs-12 col-sm-12 col-md-12 col-lg-12']")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:rateShop']")).click();
		        Thread.sleep(3000);
		         
		        String Car_val = driver.findElement(By.xpath("//*[@id='resRateLookupDlg:rateLookupForm:groups_data']")).getAttribute("innerText"); 
	            String[] temp_Car_Val = Car_val.split("\\n");
	            int len_tempCar       = temp_Car_Val.length;
	            String Excel_ValCar   = strCARGROUP;
	          
	            for (int i=1;i< len_tempCar;i++)
	            {
	               System.out.println(temp_Car_Val[i]);
	               String[] Ary_Cls_name = temp_Car_Val[i].split("-");
	               String Cls_name = Ary_Cls_name[0].trim();
	               System.out.println(Cls_name);
	               
	               if (Excel_ValCar.equals(Cls_name))
	               {
	                  System.out.println(i);
	                  int j=i-1;
	                  String Rate = "//input[@value='standard,"+j+"']";
	                  driver.findElement(By.xpath(Rate)).click();
	                  break;
	               }
	            }
                
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='resRateLookupDlg:rateLookupForm:checkOutButton']")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkuplastname']")).sendKeys(strLastName);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkupfirstname']")).sendKeys(strFirstName);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='walkUpDialog:walkUpDailogForm:walkupContinue']")).click();
		        Thread.sleep(10000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys(strDRLICCountry);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys(strDRLICState);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys(strDRLICNo);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys(strDRDOB);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys(strDRCompany);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys(strDRAddressLine1);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys(strDRAddressLine2);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys(strDRAddressLine3);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys(strDRContactInfo);
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys(strDREmailID);
		        Thread.sleep(3000);
                WebElement CheckOut_lstCOCCDC = driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType"));
                Select MOPDropDownList        = new Select(CheckOut_lstCOCCDC);
                MOPDropDownList.selectByVisibleText(strCOCCDC);
		        Thread.sleep(3000);
                driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber")).sendKeys(strCARDNO); 
		        Thread.sleep(3000);            
		        driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth")).sendKeys(strCOMM); 
		        Thread.sleep(3000);
		        driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear")).sendKeys(strCOYY); 
		        Thread.sleep(3000);            
		        WebElement CheckOut_MOPReason = driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason"));
                Select MOPReasonDropDownList  = new Select(CheckOut_MOPReason);
                MOPReasonDropDownList.selectByVisibleText(strMOPREASON);
                Thread.sleep(3000);
                driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace")).sendKeys(strCOMVA);
                Thread.sleep(3000);
                driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:mileage")).clear();
                Thread.sleep(3000);
                driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:mileage")).sendKeys(strCOMILEAGE);
                Thread.sleep(5000);
		        driver.findElement(By.id("footerForm:continueVehicleButton")).click();
		        Thread.sleep(15000);
		        driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
		        Thread.sleep(15000);
		        
		       // ########################Split Billing Flow ############################## //
		       
		       driver.findElement(By.xpath("//input[@id='footerForm:splitBillingButton']")).click();		    
		       Thread.sleep(7000);
		       
		       if(strCOCCDC.equals("American Express"))
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
		       
		       WebElement CheckOut_SB_MOP         = driver.findElement(By.id("splitBillingForm:paymentMethod"));
               Select CheckOut_SB_MOPDropDownList = new Select(CheckOut_SB_MOP);
               CheckOut_SB_MOPDropDownList.selectByVisibleText(strCO_SB_MOP);
               Thread.sleep(10000);
               WebElement CheckOut_SB_lstCOCCDC         = driver.findElement(By.id("splitBillingForm:splitBillingSwipe:ccType"));
               Select CheckOut_SB_lstCOCCDCDropDownList = new Select(CheckOut_SB_lstCOCCDC);
               CheckOut_SB_lstCOCCDCDropDownList.selectByVisibleText(strCO_SB_CCDC);
               Thread.sleep(10000);
               driver.findElement(By.id("splitBillingForm:splitBillingSwipe:ccNumber")).sendKeys(strCO_SB_CARD_NO); 
		       Thread.sleep(5000);            
		       driver.findElement(By.id("splitBillingForm:splitBillingSwipe:ccExpireMonth")).sendKeys(strCO_SB_CARD_EXP_MM); 
		       Thread.sleep(5000);
		       driver.findElement(By.id("splitBillingForm:splitBillingSwipe:ccExpireYear")).sendKeys(strCO_SB_CARD_EXP_YY); 
		       Thread.sleep(5000);            
		       WebElement CheckOut_SB_MOPReason         = driver.findElement(By.id("splitBillingForm:splitBillingSwipe:paymentReason"));
               Select CheckOut_SB_MOPReasonDropDownList = new Select(CheckOut_SB_MOPReason);
               CheckOut_SB_MOPReasonDropDownList.selectByVisibleText(strCO_SB_MOP_REASON);
               Thread.sleep(5000);
		       driver.findElement(By.id("splitBillingForm:checkoutAuth")).sendKeys("OK/12345"); 
		       Thread.sleep(5000);
		       driver.findElement(By.id("splitBillingForm:checkinAuth")).sendKeys("OK/12345"); 
		       Thread.sleep(5000);
		       
		       if (strCO_SB_DIFF_ADDRESS_OPTION.equals("Yes"))
		       {
			      driver.findElement(By.xpath("splitBillingForm:sameAsPrimaryAddress")).click();		    
			      Thread.sleep(7000);
		       }
		       else
		       {
			      driver.findElement(By.xpath("//input[@id='splitBillingForm:address1']")).sendKeys(strCO_SB_DIFF_ADDR1);
			      Thread.sleep(3000);
			      driver.findElement(By.xpath("//input[@id='splitBillingForm:address2']")).sendKeys(strCO_SB_DIFF_ADDR2);
			      Thread.sleep(3000);
			      driver.findElement(By.xpath("//input[@id='splitBillingForm:address3']")).sendKeys(strCO_SB_DIFF_ADDR3);
			      Thread.sleep(3000);
		       }
			    
			   if (strCO_SB_NO_OF_DAYS_OPTION.equals("Yes"))
			   {
				   driver.findElement(By.id("splitBillingForm:splitBillingNumberOfDays")).click();		    
				   Thread.sleep(7000);
				   String strSplitBillingTotalDaysGetText = driver.findElement(By.id("splitBillingForm:splitBillingTotalDays")).getText();
				   System.out.println(" Split Billing Total Days Value " + strSplitBillingTotalDaysGetText);
				   int strSplitBillingTotalDaysGetText2 = Integer.parseInt(strSplitBillingTotalDaysGetText);
				   int strSplitBillingTotalDaysGetText3 = strSplitBillingTotalDaysGetText2/2;
				   System.out.println(" strSplitBillingTotalDaysGetText3 value is " +strSplitBillingTotalDaysGetText3);
				   String strSplitBillingTotalDaysGetText4 =  Integer.toString(strSplitBillingTotalDaysGetText3);
				   driver.findElement(By.id("splitBillingForm:splitBillingSecondCardDays")).sendKeys(strSplitBillingTotalDaysGetText4);
				   Thread.sleep(7000);
			   }
			   //if (strCO_SB_UPSELL_OPTION.equals("Yes"))
			   //{
			        //driver.findElement(By.xpath("//input[@id='splitBillingForm:splitBillingUpsell']")).click();		    
			        //Thread.sleep(7000);
			   //}
			   //if (strCO_SB_TANDM_POINTS_OPTION.equals("Yes"))
			   //{
				   //driver.findElement(By.xpath("//input[@id='splitBillingForm:splitBillingLoyalty']")).click();		    
				   //Thread.sleep(7000);
			   //}
			   if (strCO_SB_DOLLAR_AMOUNT_OPTION.equals("Yes"))
			   {
				   driver.findElement(By.id("splitBillingForm:splitBillingDollarAmount")).click();
				   Thread.sleep(7000);
				   String strSplitBillingTotalAmountGetText = driver.findElement(By.id("splitBillingForm:splitBillingTotalAmount")).getText();
				   System.out.println(" Split Billing Total Amount value " + strSplitBillingTotalAmountGetText);
                 //String[] splitVal = strSplitBillingTotalAmountGetText.split("$");
                   String[] splitVal = strSplitBillingTotalAmountGetText.split("\\$");
                   System.out.println(splitVal[1] + " found. ");
				   double strSplitBillingTotalAmountGetText2;
				   strSplitBillingTotalAmountGetText2 = Double.parseDouble(splitVal[1]);
				   System.out.println(" strSplitBillingTotalAmountGetText2 Value is " + strSplitBillingTotalAmountGetText2);
				   double strSplitBillingTotalAmountGetText3;
				   strSplitBillingTotalAmountGetText3 = strSplitBillingTotalAmountGetText2/2;
				   System.out.println(" strSplitBillingTotalAmountGetText3 divided Value is " + strSplitBillingTotalAmountGetText3);
				   String strSplitBillingTotalAmountGetText4 = Double.toString(strSplitBillingTotalAmountGetText3);
				   System.out.println(" strSplitBillingTotalAmountGetText4 divided Value is " + strSplitBillingTotalAmountGetText4);
				   driver.findElement(By.xpath("//input[@id='splitBillingForm:splitBillingSecondCardAmount']")).click();
				   Thread.sleep(5000);
				   driver.findElement(By.xpath("//input[@id='splitBillingForm:splitBillingSecondCardAmount']")).sendKeys(strSplitBillingTotalAmountGetText4);
				   Thread.sleep(7000);
		       }
               
		       driver.findElement(By.id("splitBillingForm:splitBillingSaveButton")).click();
		       Thread.sleep(15000);
		       
		       try
		       {
		          if (driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword")).isDisplayed())
		          {
		             driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword")).sendKeys("yes");
		             Thread.sleep(10000);
			         driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
			         Thread.sleep(10000);
		          }
		       }
		       catch (Exception e)
		       {
		    	  System.out.println(" Password field is not displayed. ");
		       }
		       try
		       {
			      //driver.switchTo().alert().dismiss();
		    	  //String WindowHandle = driver.getWindowHandle();
				  //ArrayList<String> windows = new ArrayList(driver.getWindowHandles());
				  //driver.switchTo().window(windows.get(1));
		          	
		          //if (driver.findElement(By.id("printPreviewDialog:printPreviewForm:closePrintPreviewBtn")).isDisplayed())
		    	  //if (driver.findElement(By.xpath("[@id='printPreviewDialog:printPreviewForm:closePrintPreviewBtn']/span [@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only']")).isDisplayed())
		    	  //if (driver.findElement(By.id("printPreviewDialog:printPreviewForm:printPreviewDialog")).isDisplayed())
		    	  //*[@id="printPreviewDialog:printPreviewForm:printPreviewDialog"]/div[1]/a[1]/span
		    	  
		    	  Set<String> ids = driver.getWindowHandles();
		          Iterator<String> it = ids.iterator();
		          String parentid = it.next();
		          String childid = it.next();
		          driver.switchTo().window(childid);
		          
			      if (driver.findElement(By.xpath("//div[@id='printPreviewDialog:printPreviewForm:printPreviewDialog']//div//a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all']//span[@class='ui-icon ui-icon-closethick']")).isDisplayed())
		          {
			         System.out.println(" Close button is displayed. ");
		    	     //driver.findElement(By.xpath("[@id='printPreviewDialog:printPreviewForm:closePrintPreviewBtn']/span [@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only'])")).click();
		    	     //driver.findElement(By.id("printPreviewDialog:printPreviewForm:closePrintPreviewBtn")).click();
		    	     driver.findElement(By.xpath("//div[@id='printPreviewDialog:printPreviewForm:printPreviewDialog']//div//a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all']//span[@class='ui-icon ui-icon-closethick']")).click();
		    	    // driver.switchTo().alert().dismiss();
		    	     Thread.sleep(10000);
		          }
		       }
		       catch (Exception e)
		       {
		    	  System.out.println(" Close button is not displayed. ");
		       }
		       
		      // ########################Split Billing Flow ############################## //
		      
		      String strCOLNFNGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
              rwe.setCellData("INPUT_DATA", k, 81, strCOLNFNGetText);
	          String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
              rwe.setCellData("INPUT_DATA", k, 82, strCOVehicleModelGetText);
	          String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" Reservation No value is " + strCOResNoGetText);
              rwe.setCellData("INPUT_DATA", k, 83, strCOResNoGetText);
	          String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" MVA No value is " + strCOMVAGetText);
              rwe.setCellData("INPUT_DATA", k, 84, strCOMVAGetText);
	          String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" RA Number value is " + strCORANoGetText);
              rwe.setCellData("INPUT_DATA", k, 85, strCORANoGetText);
	          String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
              rwe.setCellData("INPUT_DATA", k, 86, strCOLicensePlateNoGetText);
	          String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
              rwe.setCellData("INPUT_DATA", k, 87, strCOEstimatedCostGetText);
	          String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
	          Thread.sleep(7000);
	          System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
	          rwe.setCellData("INPUT_DATA", k, 88, strCOSystemMsgGetText);
	          
	          String ScreenShotPath  = "C:\\Selenium\\Workspace\\ABG_GUI_Automation\\ScreenShots\\";
	          File file = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	          //FileUtils.copyFile(file,new File( ScreenShotPath+strTestCaseName+ ".jpg"));
	          FileUtils.copyFile(file,new File( ScreenShotPath+strTestCaseName+ ".png"));
	          
	          driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
	          Thread.sleep(7000);		    
		      
              //Log out and close tabs.
              
	          Payless_GUI_Split_Billing_CheckOut Avis = new Payless_GUI_Split_Billing_CheckOut();              
              Avis.clickLogout(driver);
              Thread.sleep(1000);
              driver.close();
              driver.quit();
	          
	          test = extent.createTest(strTestCaseName); //TestNG Extent Reports Changes.
              
			  if (rwe.getCellData("INPUT_DATA", k, 85).isEmpty())
			  {
				 rwe.setCellData("INPUT_DATA", k, 89, "FAIL");
				 test.log(Status.FAIL, "Fail"); //TestNG Extent Reports Changes.
			  }
			  else
			  {
				 test.log(Status.PASS, "Pass");
				 rwe.setCellData("INPUT_DATA", k, 89, "PASS"); //TestNG Extent Reports Changes.
			  }
	       } //End of Execute Condition: Execute.equals("Y").
        } //End of For loop.
	 }
	 finally //TestNG Extent Reports Changes.
	 {
		// TODO: handle finally clause
		extent.flush(); //TestNG Extent Reports Changes.
	 }
  }
   
  private void scriptName()
  {
	// TODO Auto-generated method stub
  }

public void clickLogout(ChromeDriver driver)
  {
	 JavascriptExecutor jse = (JavascriptExecutor)driver;
	 String clickLogoutIcon = "document.getElementById('headerLogOutButton').click()";
	 jse.executeScript(clickLogoutIcon);
	 String clickLogoutBtn = "document.getElementById('logoutForm:yesLogoutButton').click()";
	 jse.executeScript(clickLogoutBtn);
  }
} // End of AVIS_GUI_Split_Billing_CheckOut Class.