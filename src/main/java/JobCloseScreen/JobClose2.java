package JobCloseScreen;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;

import Base.Browser;

public class JobClose2 extends Browser{
	
	@FindBy(xpath = "//img[@id='ctl00_btnNew']")
	private WebElement newBtn;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$ddlService']")private WebElement selService;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$txtdate']")private WebElement tranDate;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$ddlOffice']")private WebElement office;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$ddlSerRev']")private WebElement serRevenue;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$ddlSerCost']")private WebElement costRevenue;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$txtJCJRef']")private WebElement JCJRef;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$txtGJRef']")private WebElement faVoucherRef;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$txtBookRef']")private WebElement jobRef;
	@FindBy(xpath="//textarea[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$txt_remark']")private WebElement remarks;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$JobClosing$closedJobList$btnShowDetails']")private WebElement showBtn;
	@FindBy(xpath="//img[@title='Save (Alt + Ctrl + S)']")private WebElement saveBtn;

/*	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;*/
	
	public JobClose2() {
		PageFactory.initElements(driver, this);
	}
	

	public void verifyNewBtn() throws Exception {
		Set<String> window = driver.getWindowHandles();

		Iterator<String> it = window.iterator();

		String mainpage = driver.getWindowHandle();
		while (it.hasNext()) {
			String str = it.next();
			if (!mainpage.equals(str)) {
				driver.switchTo().window(str);
			}
		}

		newBtn.click();
	}
	
	
	public void basicDetails(int excel) throws Exception {
		
		 FileInputStream file7=new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\JobClosingJrl\\TestData\\JobcloseData.xlsx");	
			
			
			XSSFWorkbook workbook=new XSSFWorkbook(file7);
			XSSFSheet sheet1 = workbook.getSheet("basicDetails");
			int rowcount = sheet1.getLastRowNum();
			int colcount = sheet1.getRow(7).getLastCellNum();
			System.out.println("basicDetails rowcount:"+rowcount+"basicDetails colcount"+colcount);

			for(int i=7;i<=rowcount;i++)
			{
				XSSFRow celldata = sheet1.getRow(i);
				try {
				int idNo = (int) celldata.getCell(1).getNumericCellValue();
				if(idNo == excel) {
		
//Select Service
					String service = celldata.getCell(2).getStringCellValue();
					Thread.sleep(1500);
					Select se1=new Select(selService);
					se1.selectByVisibleText(service);
					
//Transaction Date
					String traDate = celldata.getCell(3).getStringCellValue();
					Thread.sleep(1000);
					JavascriptExecutor js=(JavascriptExecutor)driver;
					js.executeScript("arguments[0].value='"+traDate+"'", tranDate);
					
//Office
					String office1 = celldata.getCell(4).getStringCellValue();
					Thread.sleep(1000);
					Select se3=new Select(office);
					se3.selectByVisibleText(office1);
					
//Service Revenue A/c
			/*		String serRev = celldata.getCell(5).getStringCellValue();
					Thread.sleep(1000);
					Select se4=new Select(serRevenue);
					se4.selectByVisibleText(serRev);
					
//Service Cost A/c
					String serCost = celldata.getCell(6).getStringCellValue();
					Thread.sleep(1000);
					Select se5=new Select(costRevenue);
					se5.selectByVisibleText(serCost);    
					
//JCJ Ref
					String jcj = celldata.getCell(7).getStringCellValue();
					if(JCJRef.isEnabled()) {
					Thread.sleep(1000);
					JCJRef.sendKeys(jcj);}
					
//FA Vouchar Ref
					
					String vouchar = celldata.getCell(8).getStringCellValue();
					Thread.sleep(1000);
					if(faVoucherRef.isEnabled()) {
					faVoucherRef.sendKeys(vouchar);}*/
					
//Job Ref
					
					String jobref = celldata.getCell(9).getStringCellValue();
					Thread.sleep(1000);
					JavascriptExecutor js1=(JavascriptExecutor)driver;
					js1.executeScript("arguments[0].value='"+jobref+"'", jobRef);
					
//Remarks
					String rem = celldata.getCell(10).getStringCellValue();
					Thread.sleep(1000);
					remarks.sendKeys(rem);
					
//Show Details
					Thread.sleep(1500);
					showBtn.click();

	}	
				}	
				
				catch(NullPointerException e) {
					System.out.println("Exception of loop ="+e);}			
	}
	}
		
		public void saveBtn() throws Exception{
			
			Thread.sleep(1000);
			
			saveBtn.click();
			
		}
						
}
