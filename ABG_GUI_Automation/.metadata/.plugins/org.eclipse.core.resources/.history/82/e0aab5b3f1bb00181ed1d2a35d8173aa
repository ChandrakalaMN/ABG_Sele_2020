/**
 * 
 */
package pomClasses;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author csurek
 * 27-June-2018 - POM to read value from excel
 */
public class ReadWriteExcel {

	
	public FileInputStream fis;
	public FileOutputStream fos;
	public XSSFWorkbook wb;
	public XSSFSheet sheet ;
	public XSSFCell cell;
	String XLFilePath;
	  
	public ReadWriteExcel(String XLFilePath) throws Exception
	{
		this.XLFilePath = XLFilePath;
		fis = new FileInputStream(XLFilePath);
		wb  =  new XSSFWorkbook(fis);
		fis.close();
		
	}
	
	public String getCellData(String sheetname, int row, int col)
	{
		try
		{
			sheet = wb.getSheet(sheetname);
			String data = sheet.getRow(row).getCell(col).getStringCellValue();
			return data;
		}
		catch(Exception e)
		{
			e.printStackTrace();
			return "input file exception";
		}
	}
	
	public void setCellData(String sheetname, int row, int col, String val) throws IOException
	{
		try
		{
			fis = new FileInputStream(XLFilePath);	
			wb  =  new XSSFWorkbook(fis);
		    sheet = wb.getSheet(sheetname);
		    sheet.getRow(row).getCell(col).setCellValue(val);
		    fos = new FileOutputStream(XLFilePath);  //to save file
		    wb.write(fos);
		    //wb.close();
		    fos.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
	}
	
	
	
	
}
