package AVIS.CommonFunctions;

import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

/* Customer Tab Objects */

public class CustomerObjectsFunctions {
		
	ChromeDriver driver;
	
	GUIFunctions functions = new GUIFunctions(driver);
	  
	public CustomerObjectsFunctions(WebDriver driver) throws Exception
	{
		this.driver = (ChromeDriver) driver;
		PageFactory.initElements(driver, this);
	}
	
	@FindBy(id="menulist:custInfoContainer:searchResForm:searchString")
	public  WebElement txt_displayWizard;
	
	@FindBy(id="menulist:custInfoContainer:searchResForm:searchCommandLink")
	public  WebElement btn_wizSearch;
	
	@FindBy(id="menulist:custInfoContainer:resForm:lastName")
	public  WebElement txt_custLastName;
	
	@FindBy(id="menulist:custInfoContainer:resForm:firstName")
	public  WebElement txt_custFirstName;

	@FindBy(id="menulist:custInfoContainer:resForm:wizardNumber")
	public  WebElement txt_custWizNumber;
	
	@FindBy(id="rlicenseCountry1")
	public  WebElement txt_custLicCountry;
	
	@FindBy(id="menulist:custInfoContainer:resForm:rlicenseState")
	public  WebElement txt_custLicState;
	
	@FindBy(id="menulist:custInfoContainer:resForm:licenseNumber")
	public  WebElement txt_custLicNumber;
	
	@FindBy(id="menulist:custInfoContainer:resForm:dateOfBirth")
	public  WebElement txt_custDOB;
	
	@FindBy(id="menulist:custInfoContainer:resForm:company")
	public  WebElement txt_custCompany;
	
	@FindBy(id="menulist:custInfoContainer:resForm:address1")
	public  WebElement txt_custAddr1;
	
	@FindBy(id="menulist:custInfoContainer:resForm:address2")
	public  WebElement txt_custAddr2;
	
	@FindBy(id="menulist:custInfoContainer:resForm:address3")
	public  WebElement txt_custAddr3;
	
	@FindBy(id="menulist:custInfoContainer:resForm:contactInfo")
	public  WebElement txt_custContact;
	
	@FindBy(id="menulist:custInfoContainer:resForm:rftnType")
	public  WebElement txt_custFTNType;
	
	@FindBy(id="menulist:custInfoContainer:resForm:ftNumber")
	public  WebElement txt_custFTNNo;
	
	@FindBy(id="menulist:custInfoContainer:resForm:wizconEmail")
	public  WebElement txt_custEmail;
	
	@FindBy(id="menulist:custInfoContainer:resForm:wizconEmailSpecialOffer")
	public  WebElement chk_custPromoOffer;
	
	@FindBy(id="menulist:custInfoContainer:resForm:wizconEReceipt")
	public  WebElement chk_custEReceipt;
	
	@FindBy(id="menulist:custInfoContainer:resForm:wizconVirtual")
	public  WebElement dd_custVirtual;

	@FindBy(id="menulist:custInfoContainer:resForm:carGroup")
	public  WebElement dd_custCarGroup;
	
	@FindBy(id="menulist:custInfoContainer:resForm:discountNumber")
	public  WebElement txt_custAWD;
	
	@FindBy(xpath="//input[@ng-model='custMb.res.remark' and @id='menulist:custInfoContainer:resForm:remarks']")
	public  WebElement txt_custRemarks;
	
	@FindBy(id="menulist:custInfoContainer:resForm:payMethod")
	public  WebElement dd_mop;
	
	@FindBy(id="menulist:custInfoContainer:resForm:creditCard:swipe:ccType")
	public  WebElement dd_custCreditCard;
	
	@FindBy(id="menulist:custInfoContainer:resForm:creditCard:swipe:ccNumber")
	public  WebElement txt_custCardNumber;
	
	@FindBy(id="menulist:custInfoContainer:resForm:creditCard:swipe:ccExpireMonth")
	public  WebElement txt_custCardExpiryMonth;
	
	@FindBy(id="menulist:custInfoContainer:resForm:creditCard:swipe:ccExpireYear")
	public  WebElement txt_custCardExpiryYear;
	
	@FindBy(id="footerForm:footerCustCreate")
	public  WebElement btn_custCreate;
	
	public void enterCustomerInformation(String lastname, String firstname, String wiznum, String country, String state, 
			String licnum, String dob, String company, String addr1, String addr2, String addr3, String contact, String FTNType, String FTNNumber, String email, 
			String cargroup,String awd, String remarks) throws InterruptedException
	{
	   txt_custLastName.sendKeys(lastname);	
	   txt_custFirstName.sendKeys(firstname);
       txt_custWizNumber.sendKeys(wiznum);
       txt_custLicCountry.sendKeys(country);
       txt_custLicState.sendKeys(state);
       txt_custLicNumber.sendKeys(licnum);
       txt_custDOB.sendKeys(dob);
       txt_custCompany.sendKeys(company);
       txt_custAddr1.sendKeys(addr1);
       txt_custAddr2.sendKeys(addr2);
       txt_custAddr3.sendKeys(addr3);
       txt_custContact.sendKeys(contact);
       txt_custFTNType.sendKeys(FTNType);
       txt_custFTNNo.sendKeys(FTNNumber);
       txt_custEmail.sendKeys(email);
       Select carDropDown = new Select(dd_custCarGroup);
       carDropDown.selectByVisibleText(cargroup);
       txt_custAWD.sendKeys(awd);
       txt_custRemarks.sendKeys(remarks);
	}
	
	public void enterPaymentInfo(String creditcard, String cardnumber, String expirymonth, String expiryyear)
	{
		 WebElement btn_MOPSectionExpand = driver.findElement(By.xpath("//div[@id='custPaymentInfoPanel']//div[@id='header' and @class='panel-heading'] //span[@data-target='#custPaymentInfoPanel_content' and @ng-init='custPaymentInfoPanelExpanded = false']")); 		
		 WebDriverWait wait = new WebDriverWait(driver,10);
		 wait.until(ExpectedConditions.visibilityOf(btn_MOPSectionExpand));
		 btn_MOPSectionExpand.click();
	     Select ccDropDown = new Select(dd_custCreditCard);
	     ccDropDown.selectByVisibleText(creditcard);
	     txt_custCardNumber.sendKeys(cardnumber);
	     txt_custCardExpiryMonth.sendKeys(expirymonth);
	     txt_custCardExpiryYear.sendKeys(expiryyear);	 
	}
	
	public void selectInsurance(String insval)
	{
		WebDriverWait wait = new WebDriverWait(driver, 10);
		
		if(insval.equalsIgnoreCase("LDW"))
		{
			WebElement insurance1 = driver.findElement(By.id("menulist:custInfoContainer:resForm:coverageLdwYesNo"));
			Select insLDW = new Select(insurance1);
			wait.until(ExpectedConditions.visibilityOf(insurance1));
			if (insurance1.isDisplayed())
			{
				insLDW.selectByVisibleText("Yes");
			}
		}
		else if(insval.equalsIgnoreCase("PAI"))
		{
			WebElement insurance2 = driver.findElement(By.id("menulist:custInfoContainer:resForm:coveragePaiYesNo"));
			Select insLDW = new Select(insurance2);
			wait.until(ExpectedConditions.visibilityOf(insurance2));
			if (insurance2.isDisplayed())
			{
				insLDW.selectByVisibleText("Yes");
			}
		}
		else if(insval.equalsIgnoreCase("PEP"))
		{
			WebElement insurance3 = driver.findElement(By.id("menulist:custInfoContainer:resForm:coveragePepYesNo"));
			Select insLDW = new Select(insurance3);
			wait.until(ExpectedConditions.visibilityOf(insurance3));
			if (insurance3.isDisplayed())
			{
				insLDW.selectByVisibleText("Yes");
			}
		}
		else if(insval.equalsIgnoreCase("ALI"))
		{
			WebElement insurance4 = driver.findElement(By.id("menulist:custInfoContainer:resForm:coverageAliYesNo"));
			Select insLDW = new Select(insurance4);
			wait.until(ExpectedConditions.visibilityOf(insurance4));
			if (insurance4.isDisplayed())
			{
				insLDW.selectByVisibleText("Yes");
			}
		}
		else if(insval.equalsIgnoreCase("FSO"))
		{
			WebElement insurance5 = driver.findElement(By.id("menulist:custInfoContainer:resForm:fuelServiceOption"));
			Select insLDW = new Select(insurance5);
			wait.until(ExpectedConditions.visibilityOf(insurance5));
			if (insurance5.isDisplayed())
			{
				insLDW.selectByVisibleText("Yes");
			}
		}
	}
	
	public void clickCreateBtn()
	{
		btn_custCreate.click();
	}
	
	public void searchCustomer(String wizno) throws InterruptedException
	{
		txt_displayWizard.sendKeys(wizno);
		btn_wizSearch.click();
		Thread.sleep(10000);
	}
	
	/*Modify Function*/
	public void modifyDrLicCountry(String country)
	{
		txt_custLicCountry.clear();
		txt_custLicCountry.sendKeys(country);
	}
	public void modifyDrLicState(String state)
	{
		   if(txt_custLicState.isEnabled())
		   {
			   txt_custLicState.clear();
			   txt_custLicState.sendKeys(state);
		   }
	}
	public void modifyDrLicNo(String licno)
	{
		txt_custLicNumber.clear();
		txt_custLicNumber.sendKeys(licno);
	}
	public void modifyDrCompany(String company)
	{
		txt_custCompany.clear();
		txt_custCompany.sendKeys(company);
	}
	public void modifyDrAddr1(String addr1)
	{
		txt_custAddr1.clear();
		txt_custAddr1.sendKeys(addr1);
	}
	public void modifyDrAddr2(String addr2)
	{
		txt_custAddr2.clear();
		txt_custAddr2.sendKeys(addr2);
	}
	public void modifyDrAddr3(String addr3)
	{
		txt_custAddr3.clear();
		txt_custAddr3.sendKeys(addr3);
	}
	public void modifyDrContact(String contact)
	{
		txt_custContact.clear();
		txt_custContact.sendKeys(contact);
	}
	public void modifyFTNNumber(String FTNType, String FTNNo)
	{
		txt_custFTNType.clear();
		txt_custFTNType.sendKeys(FTNType);
		txt_custFTNNo.clear();
		txt_custFTNNo.sendKeys(FTNNo);
	}
	public void modifyDrEmail(String email)
	{
		txt_custEmail.clear();
		txt_custEmail.sendKeys(email);
	}

	public void modifyPromotionalOffers(String po)
	{
        if(!chk_custPromoOffer.isSelected())
        {
        	chk_custPromoOffer.click();
        }
	}
	public void modifyEReceipt(String receipt)
	{
       if(!chk_custEReceipt.isSelected())
       {
    	   chk_custEReceipt.click();
       }
	}
	public void modifyVirtualType(String virtualType)
	{
        Select virtualTypeDD = new Select(dd_custVirtual);
        virtualTypeDD.selectByVisibleText(virtualType);
	}
	public void modifyCounterProduct(String cp)
	{
			if(cp.equalsIgnoreCase("WALL (AC)/DUAL USB 2 AMP"))
			{
                driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[1]")).click(); 
			}
			else if(cp.equalsIgnoreCase("GLASS WAIVER"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[7]")).click();   
                 System.out.println("GLASS WAIVER selected");			
			}
			else if(cp.equalsIgnoreCase("GPS UNIT"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[8]")).click(); 
				 System.out.println("GPS UNIT selected");
			}
			else if(cp.equalsIgnoreCase("MIDAS"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[11]")).click(); 
			}
			else if(cp.equalsIgnoreCase("MULTI MEDIA PLAYER"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[12]")).click(); 
			}
			else if(cp.equalsIgnoreCase("ONE-WAY GPS RATE"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[15]")).click(); 
			}
			else if(cp.equalsIgnoreCase("PARKING TICKET"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[17]")).click(); 
				 System.out.println("PARKING TICKET selected");
			}
			else if(cp.equalsIgnoreCase("ROADSIDE SAFETY NET"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[19]")).click(); 
			}
			else if(cp.equalsIgnoreCase("SAP CONSULTANT"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[20]")).click(); 
			}
			else if(cp.equalsIgnoreCase("QUERY LAN"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[21]")).click(); 
			}
			else if(cp.equalsIgnoreCase("STOP AND SHOP"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[22]")).click(); 
			}
			else if(cp.equalsIgnoreCase("STOP WATCH"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[23]")).click(); 
			}
			else if(cp.equalsIgnoreCase("PORTABLE SATELLITE RADIO"))
			{
				 driver.findElement(By.xpath("(//table[@id='menulist:custInfoContainer:resForm:wizconCPList']//input[@type='checkbox'])[25]")).click(); 
			}
			else
			{
				System.out.println("enter valid counter product to select");
			}
	}
	public void modifyCarGroup(String cargroup)
	{
	       Select carDropDown = new Select(dd_custCarGroup);
	       carDropDown.selectByVisibleText(cargroup);
	}
	public void modifyAWD(String awd)
	{
		txt_custAWD.clear();
		txt_custAWD.sendKeys(awd);
	}	
	public void modifyRemarks(String remarks)
	{
         txt_custRemarks.clear();
         txt_custRemarks.sendKeys(remarks);
	}	
	public void modifyMOP(String mop, String creditcard,String cardnumber,String expirymonth,String expiryyear) throws InterruptedException
	{
		WebElement btn_MOPSectionExpand = driver.findElement(By.xpath("//div[@id='custPaymentInfoPanel']//div[@id='header' and @class='panel-heading'] //span[@data-target='#custPaymentInfoPanel_content' and @ng-init='custPaymentInfoPanelExpanded = false']")); 		
		WebDriverWait wait = new WebDriverWait(driver,10);
		wait.until(ExpectedConditions.visibilityOf(btn_MOPSectionExpand));
		btn_MOPSectionExpand.click();	
		Select mopDropDown = new Select(dd_mop);
        String getMOPVal = mopDropDown.getFirstSelectedOption().getText();
        
        if(getMOPVal.equalsIgnoreCase(mop))
		{
			Select ccDropDown = new Select(dd_custCreditCard);
			ccDropDown.selectByVisibleText(creditcard);
			txt_custCardNumber.click();
			txt_custCardNumber.clear();
			txt_custCardNumber.sendKeys(cardnumber);
			txt_custCardExpiryMonth.click();
			txt_custCardExpiryMonth.clear();
			txt_custCardExpiryMonth.sendKeys(expirymonth);
			txt_custCardExpiryYear.click();
			txt_custCardExpiryYear.clear();
			txt_custCardExpiryYear.sendKeys(expiryyear);
		}
		else
		{
			System.out.println("MOP is not credit/debit card for this profile");
		}
	}	
	
	
//	public void captureQuickViewDetails()
//	{
//		WebElement table = driver.findElement(By.id("qvDataTable"));
//		ArrayList<WebElement> rows = (ArrayList<WebElement>) table.findElements(By.tagName("tr"));
//		for (int i1 = 1; i1 < rows.size(); i1++)
//		{
//			ArrayList<WebElement> cells = (ArrayList<WebElement>) rows.get(i1).findElements(By.tagName("td"));
//			for (int j = 0; j < cells.size(); j++)
//			{
//				String val = cells.get(j).getText();
//				if (val.isEmpty())
//				{
//					break;
//				}
//				else
//				{
//					if (j == 2)
//					{ 
//						int a = 41;
//						val = val.replaceAll("[*]", ""); // Remover * before rates
//						rwe.setCellData("INPUT_DATA", k, a, val);
//						a++;
//					}
//				}
//			}
//	    }
//	}
	
	
}
