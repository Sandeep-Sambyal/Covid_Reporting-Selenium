package Scripts;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Covid_Global {
	WebDriver driver;
	String cnfrmd,actv,recvrd,dcsd,old_cnfrmd,old_active,old_rcvrd,old_dcsd,a,b,data_tot;
	 //Constructor-------!
	
	

	 //Constructor-------!
	
	Covid_Global() {
	  
	  System.setProperty("webdriver.gecko.driver",
	  "C:\\Selenium\\JAR FIles\\geckodriver.exe"); 
	  //FirefoxOptions options = new  FirefoxOptions(); 
	  //options.setHeadless(true); 
	  driver=new  FirefoxDriver(); 
	  }
	 
	
	
	//Launch Browser-------!
	public void launch_browser() throws InterruptedException {
		
		driver.navigate().to("https://www.worldometers.info/coronavirus/");
		driver.manage().window().maximize();
		
	}
	


	//Check Values from Browser-------!
	public void fetch_values() throws IOException, InvalidFormatException, ClassNotFoundException, NoClassDefFoundError {
	
		 WebDriverWait wait = new WebDriverWait(driver,15);

		fetch_world("World");
		excel_op("World",cnfrmd, actv, recvrd, dcsd);
		fetch_data("World");
		report_data("World",cnfrmd, actv, recvrd, dcsd);
		
		t1("Canada");
		excel_op("Canada",cnfrmd, actv, recvrd, dcsd);
		fetch_data("Canada");
		report_data("Canada",cnfrmd, actv, recvrd, dcsd);
		
		t1("India");
		excel_op("India",cnfrmd, actv, recvrd, dcsd);
		fetch_data("India");
		report_data("India",cnfrmd, actv, recvrd, dcsd);
		
		//data_tot=data_tot+"</body>";
		data_tot=data_tot+"<table><br></br><br></br><br></br><td style="+"color:red"+">Note: To continue receiving this update kindly pay Rs 100/-</td></table></body>";
		System.out.println(data_tot);
		data_tot=data_tot.replace("null", "");
		fileop(data_tot);	
		send_mail(a, b);
		
		System.out.println("FREE");

	}
	
	public void fetch_world(String cntry) {
		WebDriverWait wait = new WebDriverWait(driver,15);

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table[@id='main_table_countries_today']")));
		cnfrmd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//td[text()='World']//following::td[1]")).getText();
		//System.out.println(cnfrmd);
		cnfrmd=check_string(cnfrmd);
		dcsd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//td[text()='"+cntry+"']//following::td[3]")).getText();
		dcsd=check_string(dcsd);
		recvrd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//td[text()='"+cntry+"']//following::td[5]")).getText();
		recvrd=check_string(recvrd);
		actv=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//td[text()='"+cntry+"']//following::td[6]")).getText();
		actv=check_string(actv);
		System.out.println(cntry);
		System.out.println(cnfrmd+","+actv+","+ recvrd +","+ dcsd);
//		String data="<html><body><br></br><style> table, th, td { border: 1px solid black;	border-collapse: collapse;} th,td{padding:15px;text-align:left;}</style><table border="+1+"><caption style="+"color:blue"+">PROGRESSIVE CORONA REPORT - "+cntry+"</caption>	<tr><th width="+"20%"+" colspan="+2+">Confirmed</th><th width="+"20%"+" colspan="+2+">Active</th><th width="+20+"% colspan="+2+">Recovered</th>	<th width="+20+"% colspan="+2+">Deceased</th></tr><td >"+cnfrmd+"</td><td style="+"color:red"+">Inc by "+cnfrmd+"</td><td >"+actv+"</td><td style="+"color:red"+">Inc by "+actv+"</td><td >"+recvrd+"</td><td style="+"color:red"+">Inc by "+recvrd+"</td>	<td >"+dcsd+"</td><td style="+"color:red"+">Inc by "+dcsd+"</td></tr></table>";	
//		data_tot=data_tot+data;
	}
//---- Fetch country wise values----	
	public void t1(String cntry) {
		WebDriverWait wait = new WebDriverWait(driver,15);

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table[@id='main_table_countries_today']")));
		cnfrmd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//a[text()='"+cntry+"']//following::td[1]")).getText();
		//System.out.println(cnfrmd);
		cnfrmd=check_string(cnfrmd);
		dcsd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//a[text()='"+cntry+"']//following::td[3]")).getText();
		dcsd=check_string(dcsd);
		recvrd=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//a[text()='"+cntry+"']//following::td[5]")).getText();
		recvrd=check_string(recvrd);
		actv=driver.findElement(By.xpath("//table[@id='main_table_countries_today']//a[text()='"+cntry+"']//following::td[6]")).getText();
		actv=check_string(actv);
		System.out.println(cntry);
		System.out.println(cnfrmd+","+actv+","+ recvrd +","+ dcsd);
//		String data="<html><body><br></br><style> table, th, td { border: 1px solid black;	border-collapse: collapse;} th,td{padding:15px;text-align:left;}</style><table border="+1+"><caption style="+"color:blue"+">PROGRESSIVE CORONA REPORT - "+cntry.toUpperCase()+"</caption>	<tr><th width="+"20%"+" colspan="+2+">Confirmed</th><th width="+"20%"+" colspan="+2+">Active</th><th width="+20+"% colspan="+2+">Recovered</th>	<th width="+20+"% colspan="+2+">Deceased</th></tr><td >"+cnfrmd+"</td><td style="+"color:red"+">Inc by "+cnfrmd+"</td><td >"+actv+"</td><td style="+"color:red"+">Inc by "+actv+"</td><td >"+recvrd+"</td><td style="+"color:red"+">Inc by "+recvrd+"</td>	<td >"+dcsd+"</td><td style="+"color:red"+">Inc by "+dcsd+"</td></tr></table>";	
//		data_tot=data_tot+data;
	}
	
//--- Removes comma from numbers--
	public String check_string(String a) {
		if (a.contains(",")) {
			a=a.replace(",", "");
		} 
		return a;
	}
	
	//Creates HTML Report data-------!
	public void report_data(String cntry,String confirmed,String active,String recovered,String deceased) throws IOException, InvalidFormatException, ClassNotFoundException, NoClassDefFoundError {
		String data;
		int val1= Integer.parseInt(confirmed)-Integer.parseInt(old_cnfrmd);
		int val2=Integer.parseInt(active)-Integer.parseInt(old_active);
		int val3=Integer.parseInt(recovered)-Integer.parseInt(old_rcvrd);
		int val4=Integer.parseInt(deceased)-Integer.parseInt(old_dcsd);
		System.out.println(val1+" "+val2+" "+val3+" "+val4);
		System.out.println(val1);
		//data="<body><table border="+1+"><tr >Corona Report</tr><tr><th width="+20+"%>Confirmed</th><th width="+20+"%>Active</th><th width="+20+"%>Recovered</th><th width=\"+20+\"%>Deceased</th></tr><!-- INSERT_RESULTS --><tr><td width="+15+"%>"+confirmed+"</td><td width="+15+"%>"+active+"</td><td width="+15+"%>"+recovered+"</td><td width=\"+15+\"%>"+deceased+"</td></tr></body>";
		data="<html><body><br></br><style> table, th, td { border: 1px solid black;	border-collapse: collapse;} th,td{padding:15px;text-align:left;}</style><table border=\"+1+\"><caption style="+"color:blue"+">PROGRESSIVE CORONA REPORT - "+cntry.toUpperCase()+"</caption>		<tr><th width="+"20%"+" colspan="+2+">Confirmed</th><th width="+"20%"+" colspan="+2+">Active</th><th width="+20+"% colspan="+2+">Recovered</th>	<th width="+20+"% colspan="+2+">Deceased</th></tr><td >"+confirmed+"</td><td style="+"color:red"+">Inc by "+val1+"</td><td >"+active+"</td><td style="+"color:red"+">Inc by "+val2+"</td><td >"+recovered+"</td><td style="+"color:red"+">Inc by "+val3+"</td>	<td >"+deceased+"</td><td style="+"color:red"+">Inc by "+val4+"</td></tr></table></body>";	
		data_tot=	data_tot+data;	
		//fileop(data);
		//send_mail(a,b);
		
	}
	
	//Write HTML data in txt file and then rename it to HTML-----!!!
	public void fileop(String data) throws IOException {
		File f1=new File("D:\\Data\\education\\report.txt");
		File f2=new File("D:\\Data\\education\\report.html");
		if (f2.exists()){
			f2.delete();
		}
		else if (f2.exists()){
			f2.delete();			
		}
		
		BufferedWriter out= new BufferedWriter(new FileWriter("D:\\Data\\education\\Covid reporting\\report.html"));
		out.write(data);
		out.close();
		
		f1.renameTo(f2);
		driver.close();
		
	}
	
	public void excel_op(String sheet1,String val1, String val2, String val3, String val4) throws IOException, InvalidFormatException {
		
		File file=new File("D:\\Data\\education\\Covid reporting\\RecordSheet_COVID.xlsx");
		/*boolean b=file.exists();
		if  (b==false) {
			file.createNewFile();
		}*/
		FileInputStream inputStream = new FileInputStream(file);
		XSSFWorkbook wbook= new XSSFWorkbook(inputStream);
		XSSFSheet sheet= wbook.getSheet(sheet1);
		
		
		int lastcell;
		
		
		//DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
		LocalDateTime now = LocalDateTime.now();
		String s;
		s="";
		s=s+" "+now;
		System.out.println(sheet.getLastRowNum());
		if (sheet.getLastRowNum()==-1) {
			
			lastcell=0;
			System.out.println("IN");

			 sheet.createRow(0).createCell(lastcell).setCellValue("DATE_TIME");
			 sheet.createRow(1).createCell(lastcell).setCellValue("Confirmed");
			 sheet.createRow(2).createCell(lastcell).setCellValue("Active");
			 sheet.createRow(3).createCell(lastcell).setCellValue("Recovered");
			 sheet.createRow(4).createCell(lastcell).setCellValue("Deceased");
				        
		}
		lastcell=sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
				
		sheet.getRow(0).createCell(lastcell).setCellValue(s);
		sheet.getRow(1).createCell(lastcell).setCellValue(val1);
		sheet.getRow(2).createCell(lastcell).setCellValue(val2);
		sheet.getRow(3).createCell(lastcell).setCellValue(val3);
		sheet.getRow(4).createCell(lastcell).setCellValue(val4);
		

		
		FileOutputStream fout =new FileOutputStream(file);
		wbook.write(fout);
		wbook.close();
			
	}
	
	public  void fetch_data(String sheet1) throws IOException { 
		
		File file=new File("D:\\Data\\education\\Covid reporting\\RecordSheet_COVID.xlsx");
		
		FileInputStream inputStream = new FileInputStream(file);
		XSSFWorkbook wbook= new XSSFWorkbook(inputStream);
		XSSFSheet sheet= wbook.getSheet(sheet1);
				
		int lastcell;
		lastcell=sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
		//System.out.println(lastcell);
		//System.out.println(sheet.getRow(1).getCell(lastcell-2).getCellType());
		old_cnfrmd=sheet.getRow(1).getCell(lastcell-2).getStringCellValue();
		old_active=sheet.getRow(2).getCell(lastcell-2).getStringCellValue();
		old_rcvrd=sheet.getRow(3).getCell(lastcell-2).getStringCellValue();
		old_dcsd=sheet.getRow(4).getCell(lastcell-2).getStringCellValue();
		
	
		System.out.println(old_cnfrmd+" "+old_active+" "+old_rcvrd+" "+old_dcsd);
		a=sheet.getRow(0).getCell(lastcell-2).getStringCellValue();
		b=sheet.getRow(0).getCell(lastcell-1).getStringCellValue();
		
	}
	
	
	public void send_mail(String a,String b) throws java.lang.ClassNotFoundException,java.lang.NoClassDefFoundError  {
		Properties props=new Properties();
		props.put("mail.smtp.host", "smtp.gmail.com");
		props.put("mail.smtp.socketFactory.port", "465");
		props.put("mail.smtp.socketFactory.class","javax.net.ssl.SSLSocketFactory");
		props.put("mail.smtp.auth", "true");
			props.put("mail.smtp.port", "465");
			
			Session session = Session.getDefaultInstance(props,
					 
				new Authenticator() {

					protected PasswordAuthentication getPasswordAuthentication() {

					return new PasswordAuthentication("Sambyalsin@gmail.com", "zaq12#wsx");

					}

				});

		try {
			Message message = new MimeMessage(session);
			 
			//message.setFrom(new InternetAddress("Sambyalsin@gmail.com"));
			
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("sambyalsandeep31@gmail.com"));
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("abhisinghpune11@gmail.com"));
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("sambyalpritika@gmail.com"));
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("sourabh2511991@gmail.com"));
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("sambyalgoverdhan@gmail.com"));
			message.addRecipients(Message.RecipientType.TO, InternetAddress.parse("sambyalsharda@gmail.com"));


			
			
			
	        
			message.setSubject("CORONA COUNT");

			BodyPart messageBodyPart1 = new MimeBodyPart();

			messageBodyPart1.setText("AUTOMATED MAIL:: Please find attached HTML table for progressive corona cases count in India. Table contains data from the period : "+a+" to "+b+".....");

			MimeBodyPart messageBodyPart2 = new MimeBodyPart();
			
			String filename = "D:\\Data\\education\\Covid reporting\\report.html";
			 
			// Create data source and pass the filename
			DataSource source = new FileDataSource(filename);

			// set the handler
			messageBodyPart2.setDataHandler(new DataHandler(source));

			// set the file
			messageBodyPart2.setFileName(filename);

			// Create object of MimeMultipart class
			Multipart multipart = new MimeMultipart();

			// add body part 1
			multipart.addBodyPart(messageBodyPart2);

			// add body part 2
			multipart.addBodyPart(messageBodyPart1);

			// set the content
			message.setContent(multipart);

			// finally send the email
			Transport.send(message,"Sambyalsin@gmail.com", "zaq12#wsx");

			System.out.println("=====Email Sent=====");

		} catch (MessagingException e) {

			throw new RuntimeException(e);

		}

	}
	

	
	public static void main(String[] args) throws InterruptedException   {
		// TODO Auto-generated method stub
		Covid_Global obj=new Covid_Global();
	
		
	
		try {
			obj.launch_browser();
			obj.fetch_values();
		} catch ( InvalidFormatException | ClassNotFoundException | NoClassDefFoundError | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
//		} catch (InvalidFormatException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (ClassNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (NoClassDefFoundError e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
		
		
	
		}
	}

}
