package com.cuong.automation;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import com.cuong.constan.Constant;




public class Automation {
	private static WebDriver driver;
	
	@Test(priority=0)
	public static void loginScreen() {
		System.setProperty(Constant.CHROME_DRIVER,Constant.PATH);	
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		driver.get(Constant.URL);
		driver.findElement(By.xpath(".//*[@id='login-form-username']")).click();
		driver.findElement(By.id("login-form-username")).sendKeys(Constant.USERNAME);
		driver.findElement(By.id("login-form-password")).sendKeys(Constant.PASSWORD);
		driver.findElement(By.xpath(".//*[@id='login']")).click();
		
	}
	

	@Test(priority=1)
	public void getData() {
		
		driver.findElement(By.xpath(".//*[@id='header-details-user-fullname']/span/span/img")).click();
		driver.findElement(By.id("view_profile")).click();
		String userName = driver.findElement(By.xpath(".//*[@id='details-profile-fragment']//dt[contains(text(),'Username')]")).getText();
		String valueUserName = driver.findElement(By.id("up-d-username")).getText();
		String fullName = driver.findElement(By.xpath(".//*[@id='details-profile-fragment']//dt[contains(text(),'Full name')]")).getText();
		String valueFullName= driver.findElement(By.id("up-d-fullname")).getText();
		String email= driver.findElement(By.xpath(".//*[@id='details-profile-fragment']/div[2]/ul/li/dl[4]/dt")).getText();
		String valueEmail= driver.findElement(By.xpath(".//*[@id='up-d-email']/a")).getText();
		String rememberMyLogin= driver.findElement(By.xpath(".//*[@id='details-profile-fragment']/div[2]/ul/li/dl[5]/dt")).getText();
		String valuerememberMyLogin= driver.findElement(By.id("up-d-clear-rememberme")).getText();
		String groups= driver.findElement(By.xpath(".//*[@id='details-profile-fragment']/div[2]/ul/li/dl[6]/dt")).getText();
		String valueGroup1= driver.findElement(By.xpath(".//div[@class='mod-content']//ul//dl[6]/dd[@class='description']")).getText();	
		
		String [] arrListGroup = valueGroup1.split("\n");
		String valueGroups= arrListGroup[0];
		int length= arrListGroup.length;
		for(int i=1;i<length;i++){
			 valueGroups=valueGroups+", "+ arrListGroup[i];
		}
		 
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet= workbook.createSheet("User Data");
		Map<String,Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[]{"Numerical order","Information","Detail"});
		data.put("2", new Object[]{1,userName,valueUserName});
		data.put("3", new Object[]{2,fullName,valueFullName});
		data.put("4", new Object[]{3,email,valueEmail});
		data.put("5", new Object[]{4,rememberMyLogin,valuerememberMyLogin});
	    data.put("6", new Object[]{5,groups,valueGroups});
		Set<String> keyset= data.keySet();
		int rownum = 0;
		for(String key: keyset){
			Row row = sheet.createRow(rownum++);
			Object[] objarray= data.get(key);
			 CellStyle style = workbook.createCellStyle();
			int cellnum =0;
			for(Object obj : objarray){
				Cell cell=  row.createCell(cellnum++);
				if(obj instanceof String){
					cell.setCellValue((String)obj);
				} else if(obj instanceof Integer){
					cell.setCellValue((Integer)obj);
				}
				
			}
		}
		
		File file = new File("C:/UserInformation.xls");
		if(file.exists()){
		file.delete();
		}
		FileOutputStream outFile = null;
		try {
			outFile = new FileOutputStream(file);
			workbook.write(outFile);
			System.out.println("UserInformation.xls write successfully");
		} catch (IOException e) {
			System.out.println("Error");
		}finally{
			try {
				outFile.close();
				driver.close();
			} catch (IOException e) {
				System.out.println("Error");
			}
		} 
	}
	
}
