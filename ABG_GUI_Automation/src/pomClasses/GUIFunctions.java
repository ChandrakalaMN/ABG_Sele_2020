package pomClasses;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

/*
 * @author csurek
 * 01-August-2018 - Page Factory Approach for Avis GUI objects and CR common functionalities
 */
@SuppressWarnings("unused")
public class GUIFunctions{

	ChromeDriver driver;
	  
	public GUIFunctions(WebDriver driver) throws Exception
	{
		this.driver = (ChromeDriver) driver;
		PageFactory.initElements(driver, this);
		
	}
	
	/* Login Page Objects */
	
	@FindBy(xpath="//input[@name='username']")
	private WebElement txt_userid;
	
	@FindBy(name= "PASSWORD")
	public  WebElement txt_password;

	@FindBy(xpath= "//input[@type='SUBMIT']")
    public  WebElement btn_login;

	/*Navigate TAB Objects*/	
	@FindBy(xpath="//*[@id='menubar']/ul/li[1]/a")
	private  WebElement tab_CheckOut;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[2]/a")
	private  WebElement tab_CheckIn;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[3]/a")
	private  WebElement tab_ReadyLine;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[4]/a")
	private  WebElement tab_ReservationRates;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[5]/a")
	private  WebElement tab_Manifest;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[6]/a")
	private  WebElement tab_DisplayRental;
	
	@FindBy(xpath="//*[@id='menubar']/ul/li[7]/a")
	private  WebElement tab_Customer;
	
	
	/*Create Reservation Screen Objects */
	@FindBy(id="menulist:rateshopContainer:resForm:lastName")
	public  WebElement txt_lastname;
	
	@FindBy(id="menulist:rateshopContainer:resForm:firstName")
	public  WebElement txt_firstname;
	
	@FindBy(id="menulist:rateshopContainer:resForm:checkout1_hid")
	public  WebElement dtp_COdate;
	
	@FindBy(id="menulist:rateshopContainer:resForm:checkout2")
	public  WebElement txt_COtime;
	
	@FindBy(id="menulist:rateshopContainer:resForm:returnStation:returnStation")
	public  WebElement txt_inSTA;
	
	@FindBy(id="menulist:rateshopContainer:resForm:checkin1_hid")
	public  WebElement dtp_CIdate;
	
	@FindBy(id="menulist:rateshopContainer:resForm:checkin2")
	public  WebElement txt_CItime;
	
	@FindBy(xpath="//*[@id='menulist:rateshopContainer:resForm:carGroup']")
	public  WebElement ddl_CarGroup;
	
	//expand toggle
	@FindBy(id="custToggle")
	public  WebElement btn_toggle;	
	
	/*DR LIC Details*/

	@FindBy(xpath="//*[@id='menulist:rateshopContainer:resForm:rftnType']")
	public  WebElement txt_FTNType;
	
	@FindBy(xpath="//*[@id='menulist:rateshopContainer:resForm:ftNumber']")
	public  WebElement txt_FTNNumber;
	
	@FindBy(xpath="//*[@id='menulist:rateshopContainer:resForm:discountNumber']")
	public  WebElement txt_AWD;
	
	@FindBy(id="searchCommandLinkResRateCode")
	public  WebElement btn_RateShopSearch;
	
	@FindBy(xpath="//*[@id='rateshopResDlg']/div/div/div[3]/div/div[3]/input[1]")
	public  WebElement btn_SelectRate;
	
	@FindBy(id="footerForm:footerCreateRes") 
	public  WebElement btn_CreateReservation;
	
	/*Payment Section Objects*/
	@FindBy(id="resPaymentInfoPanel") 
	public  WebElement sec_MOPSection;

	/*MOP Objects*/	
	@FindBy(id="menulist:rateshopContainer:resForm:creditCard:swipe:ccType") 
	public  WebElement dd_paymentCard;
	
	@FindBy(id="menulist:rateshopContainer:resForm:creditCard:swipe:ccNumber")
	//@FindBy(xpath="//input[@id='menulist:rateshopContainer:resForm:creditCard:swipe:ccNumber']")
	public  WebElement txt_cardNumber;
	
	@FindBy(id="menulist:rateshopContainer:resForm:creditCard:swipe:ccExpireMonth") 
	public  WebElement txt_expiryMonth;
	
	@FindBy(id="menulist:rateshopContainer:resForm:creditCard:swipe:ccExpireYear") 
	public  WebElement txt_expiryYear;
	
	@FindBy(id="menulist:rateshopContainer:resForm:creditCard:swipe:paymentReason") 
	public  WebElement dd_paymentReason;
	
	
	/*Protection&Coverage Section Objects*/
	@FindBy(id="resCoveragesInfoPanel") 
	public  WebElement sec_PAPSection;
	
	
	/*DisplayRental Objects*/
	@FindBy(id="menulist:dispormodContainer:rentalSearchForm:searchString") 
	public  WebElement txt_DRRentalSearch;
	
	@FindBy(id="searchCommandLinkDisplayRental") 
	public  WebElement btn_DRSearch;
	
	@FindBy(className ="voidRentalClass") 
	public  WebElement btn_Void;
	
	@FindBy(id="displayRentalMenuBarDialogs:voidPopupForm:voidRemarks") 
	public  WebElement txt_VoidRemarks;
	
	@FindBy(id="displayRentalMenuBarDialogs:voidPopupForm:voidPopupDlgOkButton") 
	public  WebElement txt_VoidRemarksPopupOK;
	
	/*LogOut Objects*/	
	@FindBy(id="headerLogOutButton")
	private  WebElement icon_LogOut;
	
	@FindBy(id="logoutForm:yesLogoutButton")
	private  WebElement btn_LogOut;
	
	
	/*Newly added objects*/
	/*Display Rental Screen Objects */
	@FindBy(id="menulist:dispormodContainer:rentalSearchForm:searchString")
	public  WebElement txt_displayRAResMVA;

	@FindBy(id="searchCommandLinkDisplayRental") 
	public  WebElement btn_SearchDisplayRAResMVA;

	@FindBy(xpath= "//button[@class='dropdown-toggle renExchangeMenuClass']//span[@class='caret']")
	public  WebElement btn_exchngOption;

	@FindBy(xpath="//a[@class='onTheLotClass']")
	private  WebElement elm_Onthelot;

	@FindBy(xpath="//a[@class='offTheLotClass ng-binding']")
	private  WebElement elm_Offthelot;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vxgnewmva")
	private  WebElement txtMVAnumber;	

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vxgnewmileage")
	private  WebElement txtMileage;	

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vxgnewfuelLevel")
	private  WebElement ddl_Fuel;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vehicleExgButton ui-btn")
	private  WebElement btn_onthelotok;

	@FindBy(id="rentalRepromptDialog:vehExgRepromptForm:vehExgRepromptButton")
	private  WebElement btn_onthelotenter;

	@FindBy(id="rentalRepromptDialog:vehExgRepromptForm:vehExgRepromptButton")
	private  WebElement btn_onthelotenterone;

	@FindBy(id="rentalRepromptDialog:vehExgRepromptForm:repromptValue")
	private  WebElement txtmethpay;	

	@FindBy(id="menulist:rateshopContainer:resForm:payMethod") 
	public  WebElement dd_otherMOP;
	
	/*Vehicle Exchange Objects*/
	@FindBy(id="vehicleExchange:vehicleExchangePopup:mileagein")
	public  WebElement txt_VEMileageIn;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:fuelSvc")
	private  WebElement ddl_prchsfuel;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:fuelin")
	private  WebElement ddl_fuelin;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vxgadjamt")
	public  WebElement txt_VEadjstmnt;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vehdamaged")
	private  WebElement ddl_vehdamaged;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:accidentreported")
	private  WebElement ddl_accreport;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:rentalFuelOtherType")
	private  WebElement ddl_fuelother;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:rentalFuelOtherReason")
	private  WebElement ddl_reason;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:rentalFuelOtherAmount")
	private  WebElement txt_amount;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:onewayFee")
	private  WebElement txt_oneway;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:miscCode")
	private  WebElement ddl_misc;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:miscFee")
	private  WebElement txt_miscfee;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vxgremarks")
	private  WebElement txt_VEremarks;

	@FindBy(id="vehicleExchange:vehicleExchangePopup:vehicleExgButton")
	private  WebElement btn_VEoffthelotok;

	@FindBy(id="rentalRepromptDialog:vehExgRepromptForm:vehExgRepromptButton")
	private  WebElement btn_VEoffthelotokenter;
	

    @FindBy(id="vehicleExchange:vehicleExchangePopup:delayedExchangeFlag")
    private  WebElement btn_offthelotDelayed;

    @FindBy(id="vehicleExchange:vehicleExchangePopup:exchangeDate_hid")
    private  WebElement txt_dlydInDate;

    @FindBy(id="vehicleExchange:vehicleExchangePopup:exchangeTime")
    private  WebElement txt_dlydInTime;
	

	//GUI Re-Usable Functions

	/*Open GUI Tabs*/
	public void openURL(String thinClient) throws InterruptedException, AWTException
	{
		try
		{
		driver.get(thinClient);
		Assert.assertTrue(driver.getTitle().equals("Access Manager Login")); // check whether page with title Login is loaded
		}
		catch(Exception e)
		{
			System.out.println("Login page not found");
		}
	}
	
	public void login(String username, String password) throws InterruptedException
	{
    	txt_userid.sendKeys(username);
		txt_password.sendKeys(password);
		btn_login.click();
		Thread.sleep(3000);
		String pageTitle = driver.getTitle();
	    if(pageTitle.equals("Avis Budget Group"))
		{
		  System.out.println(pageTitle);
		}
	    else
		{
		driver.navigate().refresh();
	    }
	}
	
	public void navigateToTab(String tabname)
	{
	try
		{	
			if(tabname.equalsIgnoreCase("CheckOut"))
			{
				tab_CheckOut.click();
			}
			else if(tabname.equalsIgnoreCase("CheckIn"))
			{
				tab_CheckIn.click();
			}
			else if(tabname.equalsIgnoreCase("ReadyLine"))
			{
				tab_ReadyLine.click();
			}
			else if(tabname.equalsIgnoreCase("ReservationRates"))
			{
				tab_ReservationRates.click();
			}
			else if(tabname.equalsIgnoreCase("Manifest"))
			{
				tab_Manifest.click();
			}
			else if(tabname.equalsIgnoreCase("DisplayRental"))
			{
				tab_DisplayRental.click();
			}
			else if(tabname.equalsIgnoreCase("Customer"))
			{
				tab_Customer.click();
			}
		}
		
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}		

	
	public void enterCustomerInformation(String lastname, String firstname,String codate,String cotime, String instation, String cidate, String citime) throws InterruptedException
	{
		txt_lastname.sendKeys(lastname);
		txt_firstname.sendKeys(firstname);
		dtp_COdate.sendKeys(codate);
		txt_COtime.sendKeys(cotime);
		txt_inSTA.sendKeys(instation);
		dtp_CIdate.sendKeys(cidate);
		Thread.sleep(1000);
		txt_CItime.click();
		Thread.sleep(1000);
		txt_CItime.sendKeys(citime);
		Thread.sleep(1000);
	}
	
	public void expandToggleBtn()
	{
		btn_toggle.click();
	}
	
	public void enterFTN(String FTNNumber)
	{
		String[] FTNValue = FTNNumber.split("/");
		txt_FTNType.sendKeys(FTNValue[0]);
		txt_FTNNumber.sendKeys(FTNValue[1]);
		
	}
	
	public void selectCarGroupByVT(String carGroup) throws InterruptedException
	{
	    Select carDD = new Select(ddl_CarGroup);
	    carDD.selectByVisibleText(carGroup);		
	}
	
	public void selectCarGroupByIndex(int indexno) throws InterruptedException
	{
	    Select carDD = new Select(ddl_CarGroup);
	    carDD.selectByIndex(indexno);		
	}
	
	public void clickSelectRateBtn()
	{
		btn_SelectRate.click();
	}
	
	public void enterAWD (String AWD)
	{
		txt_AWD.sendKeys(AWD);
	}
	
	/*Payment Info*/	
	public void expandPaymentInfoSection() throws InterruptedException
	{	
		
		 WebElement header_MOPPayment = sec_MOPSection.findElement(By.className("panel-heading"));

		 WebElement btn_MOPSectionExpand = header_MOPPayment.findElement(By.className("pull-right"));
		 
		 JavascriptExecutor js = (JavascriptExecutor)driver;
		 js.executeScript("arguments[0].click();", btn_MOPSectionExpand);
	     Thread.sleep(10000);
	}
	
	public void enterPaymentInformations(String cardName, String cardNumber, String month, String year, String reason) throws InterruptedException
	{
        Select cardDD = new Select(dd_paymentCard);
        cardDD.selectByVisibleText(cardName);
        txt_cardNumber.click();
        Thread.sleep(2000);
        txt_cardNumber.sendKeys(cardNumber);
        txt_expiryMonth.sendKeys(month);
        txt_expiryYear.sendKeys(year);	
        Select rsnDD = new Select(dd_paymentReason); 
        rsnDD.selectByVisibleText(reason);
	}
	
	/*Additional PRoducts*/
	
	public void expandProtectionCoverageSection()
	{	
		WebElement header_PAPPanel = sec_PAPSection.findElement(By.className("panel-heading"));

		WebElement btn_PAPSectionExpand = header_PAPPanel.findElement(By.className("pull-right")) ;
      try
       {
         if (btn_PAPSectionExpand.isDisplayed())
          {
       	  if(btn_PAPSectionExpand.isEnabled())
       	    {
       		btn_PAPSectionExpand.click();
       	    }
         } 
      }
      catch(Exception e3)
        {
       	 e3.printStackTrace();
        }
	}
	
	public void clickCreateReservationBtn()
	{
		btn_CreateReservation.click();
	}
	
	public String get_Abs_Path(){
		String Wk_dir = Paths.get("").toAbsolutePath().toString();
		return Wk_dir.toString();
	}
	
	public void create_folder_path(String Wk_dir){
		
		File path = new File(Wk_dir);
		if (!path.exists()){
			if(path.mkdirs()){
				System.out.println("Folder Created successfully");
				
			}
			else{
				
			}
			System.out.println("Folder exist!!!");
			System.out.println(path);
		}
		
	}
	
	public void ScreenCapture(String Scr_Path ,String testcasename) throws IOException
    { 
		
           Date d= new Date();
           SimpleDateFormat sdf = new SimpleDateFormat("YYYYMMdd_HH-MM-SS");
           File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);     
           FileUtils.copyFile(src,new File( Scr_Path +"\\"+testcasename+"_"+ sdf.format(d)+".png"));      
    }
	
	public void displayRA(String RANumber)
	{
		txt_DRRentalSearch.sendKeys(RANumber);
		btn_DRSearch.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}
	
	public void makeRAVoid(String voidRemarks)
	{
		btn_Void.click();
		txt_VoidRemarks.sendKeys(voidRemarks);
		txt_VoidRemarksPopupOK.click();
	}
	
	public void logout() throws InterruptedException
	{
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		icon_LogOut.click();
		Thread.sleep(2000);
		btn_LogOut.click();		
	}
	
	public void closeWindows()
	{
		ArrayList<String> tabs = new ArrayList(driver.getWindowHandles());
        driver.close(); 
		driver.quit();
	}
	
	public void searchResInCOTab(String ReservationNumber)
	{
		driver.findElement(By.id("searchString")).sendKeys(ReservationNumber);
        driver.findElement(By.id("searchCommandLink")).click();
	}
	
	
	 /*Newly added Functions*/
	
	 public void clickRateshopSearchBtn(ChromeDriver driver)
	 {
	   JavascriptExecutor jse = (JavascriptExecutor) driver;
	   String clickSearchJS = "document.getElementById('searchCommandLinkResRateCode').click()";
	   jse.executeScript(clickSearchJS);
	 }	

     public void enterSearchRAResMVAnumber(String raNumber)
     {
       txt_displayRAResMVA.sendKeys(raNumber);
     }
    
     public void enterOtherMOP(String MOP)
     {
    	Select methodDD = new Select(dd_otherMOP);
    	methodDD.selectByVisibleText(MOP);
     }
     
     public void enterAllCardPaymentInformations(String cardName, String cardNumber, String month, String year, String reason)
     {
        Select cardDD = new Select(dd_paymentCard);
        cardDD.selectByVisibleText(cardName);
        txt_cardNumber.sendKeys(cardNumber);
        Select rsnDD = new Select(dd_paymentReason); 
        if(cardName.equalsIgnoreCase("Airplus"))
        {
          rsnDD.selectByVisibleText(reason);
        }
        else
        {   
           txt_expiryMonth.sendKeys(month);
           txt_expiryYear.sendKeys(year);	
           rsnDD.selectByVisibleText(reason);
        }
      }	

     public void exchangeOptions()
     {
        if (btn_exchngOption.isDisplayed())
        {
          System.out.println("Exchange button exists");
          btn_exchngOption.click();
        } 
        else
        {
          System.out.println("Exchange button does not exist");
        }
     }
     
     public void onoffthelot(String exchngOption) throws InterruptedException
     {
       if(exchngOption.equalsIgnoreCase("On the Lot Exchange"))
       {
         Thread.sleep(3000);
         elm_Onthelot.click();
       } 
       else if(exchngOption.equalsIgnoreCase("Off the Lot Exchange"))
       {
         Thread.sleep(3000);
         elm_Offthelot.click();
       }
     }
     
     public void onthelottransaction(String MVAnumber, String Mileage, String Fuel) throws InterruptedException
     {
        txtMVAnumber.sendKeys(MVAnumber);
        Thread.sleep(2000);
        txtMileage.clear();
        txtMileage.sendKeys(Mileage);
        Select fuelDD = new Select(ddl_Fuel);
        fuelDD.selectByVisibleText(Fuel);
     }
     
     public void onthelotclickokBtn(String methpay) throws InterruptedException
     {
        btn_onthelotok.click();
        Thread.sleep(3000);
        btn_onthelotenterone.click();
        Thread.sleep(3000);
        txtmethpay.sendKeys(methpay); 
        btn_onthelotenterone.click();
        btn_onthelotenter.click();
     }

     public void notdelayedoffthelot(String mileageIn, String prchsfuel, String fuelin, String adjstmnt, String vehdamaged, String accreport, String fuelother, String reason, String amount, String oneway, String misc, String miscfee, String MVAnumber, String Mileage, String VEFuel, String remarks)  throws InterruptedException
     {
       txt_VEMileageIn.sendKeys(mileageIn);

       Select prchsfuelDD = new Select(ddl_prchsfuel);
       prchsfuelDD.selectByVisibleText(prchsfuel);

       Select fuelinDD = new Select(ddl_fuelin);
       fuelinDD.selectByVisibleText(fuelin);

       txt_VEadjstmnt.sendKeys(adjstmnt);
  
       Select vehdamagedDD = new Select(ddl_vehdamaged);
       vehdamagedDD.selectByVisibleText(vehdamaged);

       Select accreportDD = new Select(ddl_accreport);
       accreportDD.selectByVisibleText(accreport);		

       Select fuelotherDD = new Select(ddl_fuelother);
       fuelotherDD.selectByVisibleText(fuelother);

       Select reasonDD = new Select(ddl_reason);
       reasonDD.selectByVisibleText(reason);

       txt_amount.sendKeys(amount);

       txt_oneway.sendKeys(oneway);

       Select miscDD = new Select(ddl_misc);
       miscDD.selectByVisibleText(misc);

       txt_miscfee.sendKeys(miscfee);

       txtMVAnumber.sendKeys(MVAnumber);
       Thread.sleep(2000);
       txtMileage.clear();
       txtMileage.sendKeys(Mileage);

       Select fuelDD = new Select(ddl_Fuel);
       fuelDD.selectByVisibleText(VEFuel);

       txt_VEremarks.sendKeys(remarks);

       btn_VEoffthelotok.click(); 
     	    
       Thread.sleep(10000);
       btn_VEoffthelotokenter.click();
     }
     
     public void delayedoffthelot(String dlydInDate, String dlydInTime, String mileageIn, String prchsfuel, String fuelin, String adjstmnt, String vehdamaged, String accreport, String fuelother, String reason, String amount, String oneway, String misc, String miscfee, String MVAnumber, String Mileage, String VEFuel, String remarks)  throws InterruptedException
     {

       btn_offthelotDelayed.click();

       txt_dlydInDate.sendKeys(dlydInDate);

       txt_dlydInTime.sendKeys(dlydInTime);

       txt_VEMileageIn.sendKeys(mileageIn);

       Select prchsfuelDD = new Select(ddl_prchsfuel);
       prchsfuelDD.selectByVisibleText(prchsfuel);

       Select fuelinDD = new Select(ddl_fuelin);
       fuelinDD.selectByVisibleText(fuelin);

       txt_VEadjstmnt.sendKeys(adjstmnt);

       Select vehdamagedDD = new Select(ddl_vehdamaged);
       vehdamagedDD.selectByVisibleText(vehdamaged);

       Select accreportDD = new Select(ddl_accreport);
       accreportDD.selectByVisibleText(accreport);		

       Select fuelotherDD = new Select(ddl_fuelother);
       fuelotherDD.selectByVisibleText(fuelother);

       Select reasonDD = new Select(ddl_reason);
       reasonDD.selectByVisibleText(reason);

       txt_amount.sendKeys(amount);

       txt_oneway.sendKeys(oneway);

       Select miscDD = new Select(ddl_misc);
       miscDD.selectByVisibleText(misc);

       txt_miscfee.sendKeys(miscfee);

       txtMVAnumber.sendKeys(MVAnumber);
       Thread.sleep(2000);
       txtMileage.clear();
       txtMileage.sendKeys(Mileage);

       Select fuelDD = new Select(ddl_Fuel);
       fuelDD.selectByVisibleText(VEFuel);

       txt_VEremarks.sendKeys(remarks);

       btn_VEoffthelotok.click();

       Thread.sleep(10000);
       btn_VEoffthelotokenter.click();
     }
     
     /*To get RA Number*/
     public static int get_RA_Number(int RA_Num)
     {
         String str_RA_num = Integer.toString(RA_Num);
         String ary_RA     = str_RA_num.substring(8, 9);
         if (ary_RA.endsWith("6")) 
         {
                RA_Num = RA_Num + 4;
         }

         else 
         {
                RA_Num = RA_Num + 11;
         }
         return RA_Num;
  }

     
     
}
