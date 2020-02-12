package firsttestngpackage;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.Assert;
import org.testng.annotations.*;

public class FirstTestNGFile {
    public String baseUrl = "http://9653634-atwasp1/VisionControl/default.aspx";
    String driverPath = "C:\\Users\\KAMAL.RAWAL-b\\Downloads\\IEDriverServer_Win32_3.150.1\\IEDriverServer.exe";
    public WebDriver driver ; 
     
     @BeforeTest
      public void launchBrowser() {
          System.out.println("launching ie browser"); 
          System.setProperty("webdriver.ie.driver", driverPath);
          driver = new InternetExplorerDriver();
          driver.get(baseUrl);
          
      }
      @Test
      public void verifyHomepageTitle() {
    	  String expectedTitle = "Welcome: Mercury Tours";
	        String actualTitle = "";
	        actualTitle = driver.getTitle();
	        String handle= driver.getWindowHandle();
     }
      @Test(priority=0)
      public void login() throws IOException, InterruptedException, AWTException {
    	  int flag = 0;
	        int p=1;
	        int l=4;
    	  driver.findElement(By.id("tbUsername")).sendKeys("systemadministrator");
		  
		        driver.findElement(By.id("tbPassword")).sendKeys("123456789");
		      
		   
		        driver.findElement(By.id("btnOk")).click();
		        
		        
		        driver.findElement(By.id("applicationRepeater_ctl00_PermitVisionApplicationButton")).click();
		        
				   
			       
		        driver.switchTo().frame("frmVisionSuite").findElement(By.xpath("//*[@id=\"certificateKindList_ctl00_commandButton_button_MainButton\"]")).click();
		       
		        driver.findElement(By.xpath("//*[@id=\"CommandBar_rpRightMenu_ctl03_commandButton_button_MainButton\"]")).click();
		   
		      //driver.findElement(By.xpath("//*[@id=\"certificateKindList_ctl01_commandButton_button_MainButton\"]")).click();
		         //driver.findElement(By.id("certificateKindList_ctl02_commandButton_button_MainButton")).click();
		       
		        
		        
		        driver.findElement(By.xpath("//*[@id=\"ctl00_mainContent_WorkItemSearchFields_StateSearchField_StateSelector_Arrow\"]")).click();
		        driver.findElement(By.xpath("//*[@id=\"ctl00_mainContent_WorkItemSearchFields_StateSearchField_StateSelector_i0_StateCheckBox\"]")).click();
		        driver.findElement(By.xpath("//*[@id=\"ctl00_mainContent_WorkItemSearchFields_StateSearchField_StateSelector_Arrow\"]")).click();
		        
		        
		        FileInputStream fis = new FileInputStream("C:\\Users\\KAMAL.RAWAL-b\\Documents\\test.xlsx");
		        XSSFWorkbook wb = new XSSFWorkbook(fis);
		        XSSFSheet sheet = wb.getSheetAt(0);
		        int rowsNum = sheet.getPhysicalNumberOfRows();
		        
		        for(int i=0;i<rowsNum;i++)
		        {
		        	
		        	XSSFRow row = sheet.getRow(i);
			        XSSFCell cell = row.getCell(0);
			        
			        String val = String.valueOf(Math.round((cell.getNumericCellValue())));
			       
			        
			        driver.switchTo().parentFrame();
				   
			        CreateADirectory(val);
				       
		        	
		        	Thread.sleep(5000);
		        	
		        	
		            driver.switchTo().frame("frmVisionSuite").findElement(By.xpath("//*[@id=\"ctl00_mainContent_WorkItemSearchFields_RegistrationNumberSearchField_RegistrationNumberField\"]")).clear();
		        	driver.findElement(By.xpath("//*[@id=\"ctl00_mainContent_WorkItemSearchFields_RegistrationNumberSearchField_RegistrationNumberField\"]")).sendKeys(val);
		        
		       
		        driver.findElement(By.xpath("//*[@id=\"ctl00_mainContent_searchButton\"]/table/tbody/tr/td[2]")).click();
		        
		        //WebDriverWait wait2 = new WebDriverWait (driver, 40);
		        
		        Thread.sleep(5000);
		        
		        driver.switchTo().frame("ctl00_mainContent_searchResultsFrame").findElement(By.xpath("//*[@id=\"grid\"]/div[2]/div[1]/table/tbody/tr/td[4]")).click();
		        //WebDriverWait wait1 = new WebDriverWait (driver, 300);
		        
		        driver.switchTo().parentFrame();
		       
		       
		        WebElement contentFrame = driver.findElement(By.id("ctl00_WorkItemEditorPopup_popupPagePanel_pageFrame"));
		        driver.switchTo().frame(contentFrame);             
		        Thread.sleep(8000);
		        
		        driver.findElement(By.name("workItemEditor$PrintButton")).click();
		        Thread.sleep(8000);
		       
		        List<WebElement> elements = driver.findElements(By.className("printOption"));
		        System.out.println("Number of elements:" +elements.size());
                 int size = elements.size();
                 
                 
		        for (int z=3; z<elements.size();z++){
		            elements.get(z).click();
		    
                          int diff =(size - z)- 1;
                      Thread.sleep(6000);
                    Robot robot = new Robot(); 
		      //// x.sendKeys(Keys.CONTROL + "j");
               if(diff!=0)
               {
               for(int w=1;w<=diff;w++)							
               {
            	   					
            	   
            	   Thread.sleep(1000);
   	            robot.keyPress(KeyEvent.VK_TAB);
                 
               } 
               }
               
               Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	          Thread.sleep(1000);
 	            robot.keyPress(KeyEvent.VK_TAB);
 	           Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	          Thread.sleep(1000);
	            robot.keyPress(KeyEvent.VK_UP);
	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_DOWN);
  	          Thread.sleep(1000);
	            robot.keyPress(KeyEvent.VK_ENTER);
  	            
                
			if(flag==0)
               {
	                for(int y=0;y<=8;y++)
	                {
	                	Thread.sleep(1000);
	    	            robot.keyPress(KeyEvent.VK_TAB);
	                }
	                Thread.sleep(1000);
		            robot.keyPress(KeyEvent.VK_DOWN);
		            Thread.sleep(2000);
	                robot.keyPress(KeyEvent.VK_ENTER);
	                Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_TAB);
    	            Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_TAB);
    	            Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_TAB);
    	            Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_TAB);
    	            Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_TAB);
	                Thread.sleep(2000);
	                robot.keyPress(KeyEvent.VK_ENTER);
	               
               }
	                
	                
			
	               
	                	
	    	            
             
              
              if(flag==1)
              {
            	  
            	  Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            									
  	            Thread.sleep(1000);
  	            
  	            
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_RIGHT);
  	            Thread.sleep(2000);
	            robot.keyPress(KeyEvent.VK_ENTER);
  	            
  	            

              	Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
  	            Thread.sleep(1000);
  	            robot.keyPress(KeyEvent.VK_TAB);
            	  
              
            	  
                	for(int n=1;n<=p;n++)
                	{
                	Thread.sleep(1000);
    	            robot.keyPress(KeyEvent.VK_DOWN);
    	            
                	
                	}
                	 Thread.sleep(2000);
		                robot.keyPress(KeyEvent.VK_ENTER);
		                
		                
		            	Thread.sleep(1000);
		  	            robot.keyPress(KeyEvent.VK_TAB);
		  	            Thread.sleep(1000);
		  	            robot.keyPress(KeyEvent.VK_TAB);
		  	            Thread.sleep(1000);
		  	            robot.keyPress(KeyEvent.VK_TAB);
		  	            Thread.sleep(1000);
		  	            robot.keyPress(KeyEvent.VK_TAB);
		  	          Thread.sleep(1000);
		  	            robot.keyPress(KeyEvent.VK_TAB);
		                
		                
		                
		                Thread.sleep(2000);
		                robot.keyPress(KeyEvent.VK_ENTER);
                	
                	 
              }
              
             
           
		          
               
               flag=1;
		        
                
		        }
		       
		        Thread.sleep(3000);
		        
		        driver.findElement(By.id("workItemEditor_choosePrintProfilePopup_modalPopup_popup_closeButton")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.id("workItemEditor_CloseButton")).click();
		        Thread.sleep(3000);
		      p++;
		        
		        }
		        
		        
		        
	
		        
		            
		        
		        
		     
		       
		        
		        
		       

		        //close Fire fox
		        //driver.close();
		        
		       
		    }
	
	
	
	
	public static void CreateADirectory(String DirectoryName)
    {
        //project directory
        String workingDirectory = System.getProperty("user.dir");
        String  dir = workingDirectory + File.separator + DirectoryName;

        System.out.println("Final file directory : " + dir);

        //create a directory in the path

        File file = new File(dir);

        if (!file.exists()) {
            file.mkdir();
            System.out.println("Directory is created!");
        } else {
            System.out.println("Directory is already existed!");
        }

    }



      @AfterTest
      public void terminateBrowser(){
          driver.close();
      }
}