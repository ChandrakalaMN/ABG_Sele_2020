package Test_Scripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Common_ObjectRepository.ReadWriteExcel;


import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.util.IOUtils;

import Common_ObjectRepository.Repository_Zurich;


public class Amend_mutualfunds {
	com.aventstack.extentreports.ExtentReports extent;
	com.aventstack.extentreports.ExtentTest test;
	String putput1;
	@BeforeTest
	public void startReport()
	{
		extent = Common_ObjectRepository.Extentmanager.GetExtent();
	}
	
	
	@Test 
	public void login() throws Exception
{	
		
		
	
		// Read input from excel
					for (int k = 1; k <= 20; k++)
					{
						
						Amend_mutualfunds PES = new Amend_mutualfunds();
						ReadWriteExcel rwe = new ReadWriteExcel("C:\\Users\\cmn\\eclipse-workspace\\Zurich_Project\\src\\Test_Data\\Amend_Zurich_testData.xlsx");
						String Execute = rwe.getCellData("Amend", k, 1);
						
						if (Execute.equals("Y"))
						{
		
//Data properties for testcase
Properties prop = new Properties();
FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\eclipse-workspace\\Zurich_Project\\src\\Test_Data\\ZurichDatadriven.properties");
prop.load(fis);
String Test_Case = rwe.getCellData("Amend", k, 0);
String Plan_Enquiry = rwe.getCellData("Amend", k, 2);

//Delete the files in the folder
File file = new File(prop.getProperty("Screenshot"));  

String[] myFiles;    
if (file.isDirectory()) {
    myFiles = file.list();
    for (int i = 0; i < myFiles.length; i++) {
        File myFile = new File(file, myFiles[i]); 
        myFile.delete();
    }


//Screenshot path and test name
		String ScreenshotPath = prop.getProperty("Screenshot");	
		String testcasename = Test_Case;
		String xfilepath = prop.getProperty("ExcelPath") +testcasename+ ".xlsx";
		test = extent.createTest(Test_Case);
		
System.setProperty("webdriver.chrome.driver",prop.getProperty("Browser"));
ChromeDriver driver = new ChromeDriver();
Repository_Zurich functions = new Repository_Zurich(driver);
driver.manage().window().maximize();
driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);;


driver.navigate().to(prop.getProperty("URL"));
//Login using username and password and click submit
Repository_Zurich rz = new Repository_Zurich(driver);
Thread.sleep(10000);
rz.username().sendKeys(prop.getProperty("Username"));
rz.password().sendKeys(prop.getProperty("Password"));
Thread.sleep(2000);
rz.submit().click();
Thread.sleep(10000);
//driver.close();
rz.clear().click();
if (rwe.getCellData("Amend", k, 2).isEmpty()) {
	System.out.println("No Plan number Added");
} else {
	functions.Plan(Plan_Enquiry);
}

Thread.sleep(10000); 
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(2000);

/*if(rz.internalerror().isDisplayed())
{
	Thread.sleep(2000);
	String internalerr= rz.internalerror().getText();
	
	rwe.setCellData("Amend", k, 13, internalerr);
	rz.signout().click();
	driver.close();
	
	
}
else {*/
for (int i =1, a=3;i<=6 && a<=12;i++,a++)
    
{
      
String putput1= driver.findElement(By.xpath("//table[@class='pes-table pes-search-table']//tbody//tr//td[" + (i)+ "]")).getText();
                rwe.setCellData("Amend", k, a, putput1);
      
 }
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(5000);
rz.servicelink().click();
Thread.sleep(5000);
//functions.servicescreens();
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);

if(rz.amendpersonaldetails().isDisplayed())
{
rz.amendpersonaldetails().click();
Thread.sleep(10000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
}
else
{
	break;
}

if(rz.serviceservicing().isDisplayed())
{
rz.serviceservicing().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
}
else
{ break;
}

if(rz.servicinghistory().isDisplayed())
{
rz.servicinghistory().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(5000);
}

if(rz.servicesummary().isDisplayed())
{
	rz.servicesummary().click();
	Thread.sleep(5000);
	rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
	Thread.sleep(2000);
}


rz.servicecurrentpayments().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(5000);


rz.serviceplanholders().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(5000);


rz.servicetrustees().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(5000);


rz.serviceagency().click();
Thread.sleep(5000);
rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);

if(rz.servicescheme().isDisplayed())
{
	rz.servicescheme().click();
	Thread.sleep(5000);
	rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
	Thread.sleep(5000);
}

Thread.sleep(5000);

String Startdatevalue= rz.startdate().getText();
 String Valuestartdt = Startdatevalue.substring(11);
rwe.setCellData("Amend", k, 9, Valuestartdt);

String pensionplan = rz.zppp().getText();
String ppp = pensionplan.substring(29);
rwe.setCellData("Amend", k, 10, ppp);

String planholder = rz.sph().getText();
String sph = planholder.substring(11);
rwe.setCellData("Amend", k, 11, sph);

Thread.sleep(3000);
rz.returnplanentryscreen().click();
Thread.sleep(10000);

rz.ScreenCapturedate(prop.getProperty("Screenshot"), testcasename);
Thread.sleep(2000);

FileOutputStream fileOut = null;
int cntr =0;
int row = 0;
try {

       Workbook wb = new XSSFWorkbook();
      Sheet sheet = wb.createSheet("Ouput");
       // FileInputStream obtains input bytes from the image file
//FilesListFromFolder fl = new FilesListFromFolder();
       String[] pathnames;

       // Creates a new File instance by converting the given pathname string
       // into an abstract pathname
       File f = new File(prop.getProperty("Screenshot"));

       // Populates the array with names of files and directories
       pathnames = f.list();

       // For each pathname in the pathnames array
       for (String pathname : pathnames) {
             // Print the names of files and directories
             //System.out.println(pathname);

             //System.out.println("C:\\JER_Japanese\\Screenshots_Firstoption\\"+pathname);
             InputStream inputStream = new FileInputStream(
            		 prop.getProperty("Screenshot")+pathname);
             // InputStream inputStream = new
             // FileInputStream("C:\\JER_Japanese\\Screenshots\\*.png");
             // Get the contents of an InputStream as a byte[].
             //System.out.println(inputStream);
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
             //cntr = cntr+4;
             row = row +40 ;
             
             inputStream.close();
       }
       fileOut.close();

} catch (IOException ioex) {
       System.out.println(ioex);
}

if (rwe.getCellData("Amend", k, 3).isEmpty()) {
	rwe.setCellData("Amend", k, 12, "FAIL");
	test.log(Status.FAIL, "Fail");
} else {
	rwe.setCellData("Amend", k, 12, "PASS");
	test.log(Status.PASS, "Pass");
}

Thread.sleep(5000);

//rz.signout().click();
//driver.close();
}
						}//End of IF line 74
					
						//rz.signout().click();
						//driver.close();					
}///End of FOR line 64
					
}

}
	
	 

//}






					

					


	





