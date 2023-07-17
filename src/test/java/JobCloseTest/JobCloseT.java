package JobCloseTest;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import Base.Browser;
import JobCloseScreen.JobClose1;
import JobCloseScreen.JobClose2;

public class JobCloseT extends Browser {
	JobClose1 jb1;
	JobClose2 jb2;
	
	@BeforeMethod()
	public void setup() throws Exception {
		
		initilization();
		jb1=new JobClose1();
		jb2=new JobClose2();
		Thread.sleep(1000);
		jb1.verifyLoginApp();
		jb1.verifyIFFBtn();
		Thread.sleep(2000);
		jb1.verifyFinanceBtn();
		Thread.sleep(2000);
		jb1.verifyJobCloseBtn();
		Thread.sleep(2000);
		jb2.verifyNewBtn();
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}
	
	@Test()
	public void data() throws Exception {
		
		FileInputStream file1=new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\JobClosingJrl\\TestData\\JobcloseData.xlsx");	
	
		XSSFWorkbook workbook=new XSSFWorkbook(file1);
		XSSFSheet sheet = workbook.getSheet("basicDetails");
		int rowcount = sheet.getLastRowNum();
		int row= rowcount - 6;
		int colcount = sheet.getRow(7).getLastCellNum();
		System.out.println("rowcount in test:"+row+" colcount in test:"+colcount);
	
	for(int exec=1;exec<=row;exec++) {
		Thread.sleep(2000);
	
		jb2.basicDetails(exec);
		jb2.saveBtn();
		System.out.println("*** JOB CLOSING DONE : "+exec+" ***");

}
		
	}
	
	@AfterMethod()
	public void exit() {
		driver.close();
	}

}
