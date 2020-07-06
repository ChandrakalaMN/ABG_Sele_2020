package pomClasses;

import java.text.SimpleDateFormat;
import java.util.Calendar;

//import org.testng.Assert;
//import org.testng.ITestResult;
//import org.testng.SkipException;
//import org.testng.annotations.AfterMethod;
//import org.testng.annotations.AfterTest;
//import org.testng.annotations.BeforeTest;
//import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class Extentmanager {

	private static ExtentReports extent;
	private static ExtentTest test;
	private static ExtentHtmlReporter htmlReporter;
	private static String filePath = "C:/Report/GUI_Report_"+new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime())+".html";

	public static ExtentReports GetExtent() {
		if (extent != null)
			return extent; // avoid creating new instance of html file
		extent = new ExtentReports();
		extent.attachReporter(getHtmlReporter());
		return extent;
	}

	private static ExtentHtmlReporter getHtmlReporter() {
		htmlReporter = new ExtentHtmlReporter(filePath);
		htmlReporter.loadXMLConfig("C:/Report/extent-config.xml");
		return htmlReporter;
	}

	public static ExtentTest createTest(String name, String description) {
		//extent = new ExtentReports (System.getProperty("user.dir") +"/test-output/STMExtentReport.html", true);
		test = extent.createTest(name, description);
		return test;
	}
} 