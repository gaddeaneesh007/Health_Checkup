package com.dell.healthcheckup.daily;

import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import java.util.concurrent.TimeUnit;

import org.apache.commons.mail.DefaultAuthenticator;

import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.MultiPartEmail;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Health_Checkup {
	static WebDriver driver;
	static ArrayList<String> Reagion = new ArrayList<String>();
	static ArrayList<String> Source = new ArrayList<String>();
	static ArrayList<String> OrderNumber = new ArrayList<String>();
	static ArrayList<String> OrderStatus = new ArrayList<String>();
	public static void main(String[] args) throws InterruptedException, IOException, EmailException {
		DriverandReadexcel di = new DriverandReadexcel();
		String Smessage = null;
		driver=di.Initialization();
		//WebDriverWait wait = new WebDriverWait(driver, 10000);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);	
		di.openSite();
		Select Region = new Select(driver.findElement(By.xpath("//Select[@id='ddlRegion']")));
		Select Src_Env = new Select(driver.findElement(By.xpath("//Select[@id='srcEnv']")));
		Select Des_Env = new Select(driver.findElement(By.xpath("//Select[@id='destEnv']")));
		WebElement order_Number= driver.findElement(By.xpath("//input[@id='srcOrder']"));
		WebElement create_order=driver.findElement(By.xpath("//input[@id='btnCreate']"));
		String path = System.getProperty("user.dir")+"/TestData/Checkup.xlsx";
		
		System.out.println(path);
		
		String ordernumber=null;
		
		File file=new File(path);
		FileInputStream inputStream = new FileInputStream(file);	
		XSSFWorkbook wb=new XSSFWorkbook(inputStream);
		XSSFSheet sheet= wb.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		System.out.println(lastRow);
		String regionName=null,sourceEnv=null,DesEnv=null,oldOrdernumber=null,newordernumber=null,currentStatus=null;
		int id=0;
		
		for(int i=1;i<=lastRow;i++)
		{
			regionName=sheet.getRow(i).getCell(0).getStringCellValue();
			Reagion.add(regionName);
			sourceEnv=sheet.getRow(i).getCell(1).getStringCellValue();
			DesEnv=sheet.getRow(i).getCell(2).getStringCellValue();
			Source.add(DesEnv);
			id = (int)sheet.getRow(i).getCell(3).getNumericCellValue();
			oldOrdernumber = String.valueOf(id);
			System.out.println(regionName);
			Region.selectByVisibleText(regionName);
			Src_Env.selectByVisibleText(sourceEnv);
			Des_Env.selectByVisibleText(DesEnv);
			order_Number.clear();
			order_Number.sendKeys(oldOrdernumber);
			order_Number.sendKeys(Keys.TAB);
			create_order.click();
			Thread.sleep(5000);
			
			if(driver.findElements(By.xpath("//div[@id='divResultTable2']//td[4]")).size()>0)
			{
			//driver.findElement(By.xpath("//input[@id='btnRefresh']")).click();
			Thread.sleep(2000);
			newordernumber= driver.findElement(By.xpath("//div[@id='divResultTable2']//td[4]")).getText();
			OrderNumber.add(newordernumber);
			sheet.getRow(i).createCell(4).setCellValue(newordernumber);
			
			}
			else
			{
				OrderNumber.add("Order Not Created");
				sheet.getRow(i).createCell(4).setCellValue("Order Not Created");
			}
			
		}
		
		for(int i=1;i<=lastRow;i++)
		{
			regionName=sheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println("regionName:-"+regionName);
			DesEnv=sheet.getRow(i).getCell(2).getStringCellValue();
			newordernumber=sheet.getRow(i).getCell(4).getStringCellValue();
			if(!newordernumber.equalsIgnoreCase("Order Not Created")){
			currentStatus=di.openOfsSite(regionName,DesEnv,newordernumber);	
			OrderStatus.add(currentStatus);
			System.out.println("Current Ststaus Recived"+currentStatus);
			sheet.getRow(i).createCell(5).setCellValue(currentStatus);
			}
			else{
				OrderStatus.add(currentStatus);
				sheet.getRow(i).createCell(5).setCellValue("Order Not Created");
			}
			
			
		}
		FileOutputStream fout = new FileOutputStream(file);
		wb.write(fout);
		fout.close();
		wb.close();
		System.out.println("Workbook closed");
		System.out.println("End of the program");
		String html=CreateHTMLfile();
		sendEmail(html);
		}
	public static String CreateHTMLfile()
	{
		StringBuilder buf = new StringBuilder();
		buf.append("<table>" +
		           "<tr>" +
		           "<th>Reagion</th>" +
		           "<th>Order Number</th>" +
		           "<th>Order Status</th>" +
		           "</tr>");
		for (int i = 0; i < Reagion.size(); i++) {
		    buf.append("<tr><td>")
		       .append(Reagion.get(i))
		       .append("</td><td>")
		       .append(OrderNumber.get(i))
		       .append("</td><td>")
		       .append(OrderStatus.get(i))
		       .append("</td></tr>");
		}
		buf.append("</table>");
		String html = buf.toString();
		System.out.println(html);
		return html;
	}
	
		public static void sendEmail(String tableData) throws EmailException
		{
			// Create the attachment
			  EmailAttachment attachment = new EmailAttachment();
			  String aPath = System.getProperty("user.dir")+"/TestData/Checkup.xlsx";
			  attachment.setPath(aPath);
			  attachment.setDisposition(EmailAttachment.ATTACHMENT);
			  attachment.setDescription("Daily Health Checkup");
			  attachment.setName("Health_CheckUp");
			// Create the email message
			MultiPartEmail email = new MultiPartEmail();
			
			email.setHostName("smtp.googlemail.com");
			email.setSmtpPort(587);
			email.setAuthenticator(new DefaultAuthenticator("dhoni885883@gmail.com", "satish12$"));
			email.setSSLOnConnect(true);
			email.addTo("Aneesh_Gadde@DellTeam.com", "Aneesh Gadde");
			email.setFrom("dhoni885883@gmail.com", "Me");
			DateFormat df = new SimpleDateFormat("dd/MM/yy HH:mm:ss");
		    Date dateobj = new Date();
		    System.out.println(df.format(dateobj));
		    email.setSubject("Health Check up on "+dateobj);
			email.setMsg(tableData);
			 
			// add the attachment
			email.attach(attachment);
			 
			// send the email
			email.send();
			System.out.println("Mail sent Sucessfully");
	
		}
		
	}


