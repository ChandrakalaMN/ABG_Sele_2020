/**
'#############################################################################################################################
'## SCRIPT NAME:                           Budget_GUI_ReadyLine_Regression
'## BRAND:                                 Budget
'## DESCRIPTION:                           Regress the Readyline screen.
'## FUNCTIONAL AREA :                      Readyline
'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
'## OUTPUT:                                Readyline reports should be retrieved successfully.
'#############################################################################################################################
 **/

package Budget.TestScripts;
import java.io.*;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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

public class Budget_GUI_ReadyLine_Regression
{
   
	ExtentReports extent; //TestNG Extent Reports Changes.
	ExtentTest test;     //TestNG Extent Reports Changes.
	
	@BeforeTest  //TestNG Extent Reports Changes.
	public void startReport() //TestNG Extent Reports Changes.
	{
		//extent = Extentmanager.GetExtent(this.getClass().getSimpleName()); //TestNG Extent Reports Changes.
		extent = Extentmanager.GetExtent();
	}
	   
	//public static void main(String[] args) throws IOException, Exception, FileNotFoundException
	
	@Test //TestNG Extent Reports Changes.
	public void test() throws Exception //TestNG Extent Reports Changes.
	{
	   try //TestNG Extent Reports Changes.
	   {
	      // ead input from excel
	     // need to find the no of iterations / iterations count.
	     
	     String Temp = "Location";
	     
	     ChromeDriver driver = null;
	     //int intRowCount = 0;
	     
	     ReadWriteExcel rwe  = new ReadWriteExcel("C:\\GitHub\\ABG\\Budget\\TestData\\InternetGUI\\Budget_GUITestData_ReadyLine_Regression.xlsx");
         
	     int intRowCount = rwe.sheet.getLastRowNum();
	     
	     System.out.println(" total row count in ReadyLine test data sheet is : " + intRowCount);
	     
	     for (int k=1;k<=intRowCount;k++)
	     {
            String Execute = rwe.getCellData("INPUT_DATA",k,2);
            
		    if (Execute.equals("Y"))
            {
			   
			    String outSTA    = rwe.getCellData("INPUT_DATA", k, 10);
			    
			    System.out.println(" Line No. 70. Oustation value is : " + outSTA);
			    System.out.println(" Line No. 71. Temp Location value is : " + Temp);
			    
			    if (!Temp.equalsIgnoreCase(outSTA))
	            {
				    System.out.println(" Line No. 75. Inside if : Temp Outstation value is : " + Temp);
				    System.setProperty("webdriver.chrome.driver","C:\\Program Files (x86)\\Selenium\\WebDrivers\\ChromeDrivers\\chromedriver.exe");
			        driver = new ChromeDriver();
			        driver.manage().window().maximize();
			        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
     		        System.out.println(" Current iteration : " + k);			
				    String clientURL = rwe.getCellData("INPUT_DATA", k, 7);
				    
			        try
			        {
			           String thinClient = clientURL+outSTA;
			           driver.get(thinClient);
			           Assert.assertTrue(driver.getTitle().contains("Login")); // check whether page with title Login is loaded
			        }
			        catch(Exception e)
			        {
			           System.out.println("Login page not found");
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
                   
	    	      //click ReadyLine tab.
			      
			      driver.findElement(By.xpath("//*[@id='menubar']/ul/li[3]/a")).click();
			      Thread.sleep(500);
			      
	           } //Location compare condition !Temp.equalsIgnoreCase(outSTA) Ends.
			   
			   Temp = outSTA;
			   
		      //Reading values from Excel.
			  
     	      String strOption      = rwe.getCellData("INPUT_DATA", k, 11);     	    
/*     	      String strMVA1        = rwe.getCellData("INPUT_DATA", k, 12);
		    
		    if (!strMVA1.equals(null))
		    {  
		       //Inlcude script for Add cars to readyline.
		       
		       System.out.println(" strMVA1 is not empty. ");

		       String strMVA2        = rwe.getCellData("INPUT_DATA", k, 13);
		       String strMVA3        = rwe.getCellData("INPUT_DATA", k, 14);
		       String strMVA4        = rwe.getCellData("INPUT_DATA", k, 15);
		       String strMVA5        = rwe.getCellData("INPUT_DATA", k, 16);
		       String strMVA6        = rwe.getCellData("INPUT_DATA", k, 17);
		       String strMVA7        = rwe.getCellData("INPUT_DATA", k, 18);
		       String strMVA8        = rwe.getCellData("INPUT_DATA", k, 19);
		       String strMVA9        = rwe.getCellData("INPUT_DATA", k, 20);
		       String strMVA10       = rwe.getCellData("INPUT_DATA", k, 21);
		       String strMileage1    = rwe.getCellData("INPUT_DATA", k, 22);
		       String strMileage2    = rwe.getCellData("INPUT_DATA", k, 23);
		       String strMileage3    = rwe.getCellData("INPUT_DATA", k, 24);
		       String strMileage4    = rwe.getCellData("INPUT_DATA", k, 25);
		       String strMileage5    = rwe.getCellData("INPUT_DATA", k, 26);
		       String strMileage6    = rwe.getCellData("INPUT_DATA", k, 27);
		       String strMileage7    = rwe.getCellData("INPUT_DATA", k, 28);
		       String strMileage8    = rwe.getCellData("INPUT_DATA", k, 29);
		       String strMileage9    = rwe.getCellData("INPUT_DATA", k, 30);
		       String strMileage10   = rwe.getCellData("INPUT_DATA", k, 31);
		       String strSpace1      = rwe.getCellData("INPUT_DATA", k, 32);
		       String strSpace2      = rwe.getCellData("INPUT_DATA", k, 33);
		       String strSpace3      = rwe.getCellData("INPUT_DATA", k, 34);
		       String strSpace4      = rwe.getCellData("INPUT_DATA", k, 35);
		       String strSpace5      = rwe.getCellData("INPUT_DATA", k, 36);
		       String strSpace6      = rwe.getCellData("INPUT_DATA", k, 37);
		       String strSpace7      = rwe.getCellData("INPUT_DATA", k, 38);
		       String strSpace8      = rwe.getCellData("INPUT_DATA", k, 39);
		       String strSpace9      = rwe.getCellData("INPUT_DATA", k, 40);
		       String strSpace10     = rwe.getCellData("INPUT_DATA", k, 41);
		       String strFuel1       = rwe.getCellData("INPUT_DATA", k, 42);
		       String strFuel2       = rwe.getCellData("INPUT_DATA", k, 43);
		       String strFuel3       = rwe.getCellData("INPUT_DATA", k, 44);
		       String strFuel4       = rwe.getCellData("INPUT_DATA", k, 45);
		       String strFuel5       = rwe.getCellData("INPUT_DATA", k, 46);
		       String strFuel6       = rwe.getCellData("INPUT_DATA", k, 47);
		       String strFuel7       = rwe.getCellData("INPUT_DATA", k, 48);
		       String strFuel8       = rwe.getCellData("INPUT_DATA", k, 49);
		       String strFuel9       = rwe.getCellData("INPUT_DATA", k, 50);
		       String strFuel10      = rwe.getCellData("INPUT_DATA", k, 51);
		       String strReadyLine1  = rwe.getCellData("INPUT_DATA", k, 52);
		       String strReadyLine2  = rwe.getCellData("INPUT_DATA", k, 53);
		       String strReadyLine3  = rwe.getCellData("INPUT_DATA", k, 54);
		       String strReadyLine4  = rwe.getCellData("INPUT_DATA", k, 55);
		       String strReadyLine5  = rwe.getCellData("INPUT_DATA", k, 56);
		       String strReadyLine6  = rwe.getCellData("INPUT_DATA", k, 57);
		       String strReadyLine7  = rwe.getCellData("INPUT_DATA", k, 58);
		       String strReadyLine8  = rwe.getCellData("INPUT_DATA", k, 59);
		       String strReadyLine9  = rwe.getCellData("INPUT_DATA", k, 60);
		       String strReadyLine10 = rwe.getCellData("INPUT_DATA", k, 61);
		       
		       driver.findElement(By.id("menulist:readyLineSubMenu\"]/div/div/div/div[7]/button")).click();
		       //driver.findElement(By.id("menulist:readyLineSubMenu")).click();
		     //*[@id="menulist:readyLineSubMenu"]/div/div/div/div[1]/div/button/span
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-0")).sendKeys(strMVA1);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-0")).sendKeys(strMileage1);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-0")).sendKeys(strSpace1);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-0")).sendKeys(strFuel1);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-0")).sendKeys(strReadyLine1);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-1")).sendKeys(strMVA2);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-1")).sendKeys(strMileage2);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-1")).sendKeys(strSpace2);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-1")).sendKeys(strFuel2);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-1")).sendKeys(strReadyLine2);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-2")).sendKeys(strMVA3);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-2")).sendKeys(strMileage3);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-2")).sendKeys(strSpace3);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-2")).sendKeys(strFuel3);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-2")).sendKeys(strReadyLine3);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-3")).sendKeys(strMVA4);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-3")).sendKeys(strMileage4);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-3")).sendKeys(strSpace4);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-3")).sendKeys(strFuel4);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-3")).sendKeys(strReadyLine4);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-4")).sendKeys(strMVA5);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-4")).sendKeys(strMileage5);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-4")).sendKeys(strSpace5);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-4")).sendKeys(strFuel5);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-4")).sendKeys(strReadyLine5);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-5")).sendKeys(strMVA6);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-5")).sendKeys(strMileage6);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-5")).sendKeys(strSpace6);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-5")).sendKeys(strFuel6);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-5")).sendKeys(strReadyLine6);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-6")).sendKeys(strMVA7);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-6")).sendKeys(strMileage7);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-6")).sendKeys(strSpace7);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-6")).sendKeys(strFuel7);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-6")).sendKeys(strReadyLine7);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-7")).sendKeys(strMVA8);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-7")).sendKeys(strMileage8);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-7")).sendKeys(strSpace8);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-7")).sendKeys(strFuel8);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-7")).sendKeys(strReadyLine8);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-8")).sendKeys(strMVA9);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-8")).sendKeys(strMileage9);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-8")).sendKeys(strSpace9);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-8")).sendKeys(strFuel9);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-8")).sendKeys(strReadyLine9);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mva-9")).sendKeys(strMVA10);
		       Thread.sleep(5000);
		       driver.findElement(By.id("mileage-9")).sendKeys(strMileage10);
		       Thread.sleep(5000);
		       driver.findElement(By.id("space-9")).sendKeys(strSpace10);
		       Thread.sleep(5000);
		       driver.findElement(By.id("fuel-9")).sendKeys(strFuel10);
		       Thread.sleep(5000);
		       driver.findElement(By.id("readyline-9")).sendKeys(strReadyLine10);
		       Thread.sleep(5000);
		       driver.findElement(By.id("addButton")).click();
		       Thread.sleep(5000);
		    }
		    else
		    {
		    	System.out.println(" strMVA1 is empty. ");
		    }
		    
		    String strVTMVA = rwe.getCellData("INPUT_DATA", k, 62);
		    
		    if (!strVTMVA.equals(null))
		    {
			    //Inlcude script for Vehicle Trace.
		    	 
			    System.out.println(" strVTMVA is not empty. ");
		    	
		        String strVTFromDate = rwe.getCellData("INPUT_DATA", k, 63);
		        
			    driver.findElement(By.id("menulist:readyLineSubMenu\"]/div/div/div/div[8]/button")).click();
			    Thread.sleep(5000);
			    driver.findElement(By.id("vehTracedialog:vehTraceForm:mva")).sendKeys(strVTMVA);
			    Thread.sleep(5000);
			    driver.findElement(By.id("vehTracedialog:vehTraceForm:fromDate_hid")).sendKeys(strVTFromDate);
			    Thread.sleep(5000);
			    driver.findElement(By.id("vehTracedialog:vehTraceForm:vehTraceCmdBtn")).click();
			    Thread.sleep(5000);
			    driver.findElement(By.id("vehTracedialog:vehTraceForm:closeCmdBtn")).click();
			    Thread.sleep(5000);
		    }
		    else
		    {
		    	System.out.println(" strVTMVA is empty. ");
		    }*/
            
			Thread.sleep(500);
			
            //Select ReadyLine Type.
            
			System.out.println(" Ready Line Type value is : " + strOption);
			driver.findElement(By.xpath("//*[@id=\'menulist:readyLineSubMenu\']/div/div/div/div[1]/div/button")).click();
			driver.findElement(By.linkText(strOption)).click();
			Thread.sleep(7000);
			
		    WebElement RLitemCount = driver.findElement(By.className("ui-grid-pager-count-container"));
		    
		    Thread.sleep(5000);
		    
		    if(RLitemCount.findElement(By.xpath("//*[@id=\'ll_showReadylineGrid\']/div/div[2]/div[2]/div/span")).isDisplayed())
		    {
                String strRLCountVal = RLitemCount.findElement(By.xpath("//*[@id=\'ll_showReadylineGrid\']/div/div[2]/div[2]/div/span")).getText();
                System.out.println("Ready Line Count is : " +strRLCountVal);
                String[] splitVal = strRLCountVal.split("of");
                System.out.println(splitVal[1]+ " found. ");
                rwe.setCellData("INPUT_DATA", k, 74, splitVal[1]+ " found. ");
		    }
            else
            {
                System.out.println(" 0 items found. ");
                rwe.setCellData("INPUT_DATA", k, 74, " 0 items found. ");
		    }

	          String strTestCaseName = rwe.getCellData("INPUT_DATA", k, 4);
	          test = extent.createTest(strTestCaseName); //TestNG Extent Reports Changes.
              
			  if (rwe.getCellData("INPUT_DATA", k, 74).isEmpty())
			  {
				 rwe.setCellData("INPUT_DATA", k, 77, "PASS");
				 test.log(Status.PASS, "Pass"); //TestNG Extent Reports Changes.
			  }
			  else
			  {
				 test.log(Status.PASS, "Pass");
				 rwe.setCellData("INPUT_DATA", k, 77, "PASS"); //TestNG Extent Reports Changes.
			  }
           }
         } //End of For loop.
	     
	     Budget_GUI_ReadyLine_Regression Avis = new Budget_GUI_ReadyLine_Regression();
         
         Avis.clickLogout(driver);
         Thread.sleep(1000);
         driver.close();
         driver.quit();
	  }
	  finally //TestNG Extent Reports Changes.
	  {
		 // TODO: handle finally clause
		 extent.flush(); //TestNG Extent Reports Changes.
	  }
   } // End of public static void main.
   
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
   
} // End of Budget_GUI_ReadyLine_Regression Class.