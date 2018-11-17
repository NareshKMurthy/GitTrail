package Generics;

import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Utility {
	
	
	public static String getFormatedDateTime1(){
		SimpleDateFormat simpleDate = new SimpleDateFormat("dd_MM_yyyy_hh_mm_ss");
		return simpleDate.format(new Date());
	}
	
	public static String getScreenShot(WebDriver driver, String imageFolderPath){
		String imagePath=imageFolderPath+"/"+getFormatedDateTime1()+".png";
		TakesScreenshot page=(TakesScreenshot) driver;
		try{
			FileUtils.copyFile(page.getScreenshotAs(OutputType.FILE), new File(imagePath));
		}catch(Exception e){
			
		}
		return imagePath;
		
	}
	
	public static String getScreenShot(String imageFolderPath){
		String imagePath=imageFolderPath+"/"+getFormatedDateTime1()+".png";
		
		try{
			Dimension desktopSize=Toolkit.getDefaultToolkit().getScreenSize();
			BufferedImage image = new Robot().createScreenCapture(new Rectangle(desktopSize));
			ImageIO.write(image, "png", new File(imagePath));
		}
		catch(Exception e){
		}

		return imagePath;
		
	}

	public String getTagValueFromXMLFile(String XMLFileName, String tag) throws Exception{
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.parse(XMLFileName);
		doc.getDocumentElement().normalize();
		System.out.println("Root element : "+ doc.getDocumentElement().getNodeName());	
		String tagValue = null;	
		NodeList nodelist = doc.getElementsByTagName("*");
		for(int i=0;i<nodelist.getLength();i++){
			Element element = (Element) nodelist.item(i);
			if(element.getNodeName().equals(tag)){
//				System.out.println("Node is "+element.getNodeName() + " and the value is : "+element.getTextContent());
//				tagValue = ((Object) element).getTextContent();
				break;
			}
		}
		return tagValue;
	}
	
	public void putCellData(String fileName, String sheetName, int row, int col, String data) throws Exception{
		FileInputStream fis = new FileInputStream(fileName);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet s = wb.getSheet(sheetName);
		try{
			Row r = s.getRow(row);
			Cell c = r.getCell(col);
			if(c == null){
				c = r.createCell(col);
				c.setCellValue(data);
			}
			else{
				c.setCellValue(data);
			}
			FileOutputStream fos = new FileOutputStream(fileName);
			wb.write(fos);
					
		}
		catch(Exception e){
			e.printStackTrace();
		}	
	}	
	  
	public static void selectFromDropdown(WebElement listBox, String month) {
		Select select = new Select(listBox);
		select.selectByVisibleText(month);
	}
	
	public void failTestCaseReport(WebDriver driver, ExtentTest eTest, String SNAP_PATH){
		String imgPath = Utility.getScreenShot(driver,SNAP_PATH);
		String path = eTest.addScreenCapture("."+imgPath);
		eTest.log(LogStatus.FAIL,"Script failed",path);
	}
	
	public void passTestCaseReport(WebDriver driver, ExtentTest eTest, String SNAP_PATH){
		String imgPath = Utility.getScreenShot(driver,SNAP_PATH);
		String path = eTest.addScreenCapture("."+imgPath);
		eTest.log(LogStatus.PASS,"Script executed successfully",path);
	}
	
}
