package PharmaSunday;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ApolloPharma_thdhalf{

	
		public static void main(String[] args) {
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--headless"); // Run Chrome in headless mode
			options.addArguments("--disable-gpu"); // Disable GPU acceleration
			options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			options.addArguments("--start-maximized");	

			ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

			// Schedule the task to run every day at 7:00 AM
			Calendar now = Calendar.getInstance();
			Calendar nextRunTime = Calendar.getInstance();
			nextRunTime.set(Calendar.HOUR_OF_DAY, 16);
			nextRunTime.set(Calendar.MINUTE, 0);
			nextRunTime.set(Calendar.SECOND, 0);

			long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
			if (initialDelay < 0) {
				initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
			}

			scheduler.scheduleAtFixedRate(() -> {
				try {
					System.out.println("Starting web scraping task...");
					ApolloPharma.runWebScraping();
					System.out.println("Web scraping task completed.");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
		}
		public static void runWebScraping() throws Exception{

			//      System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
			WebDriver driver = new ChromeDriver();
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--headless"); // Run Chrome in headless mode
			options.addArguments("--disable-gpu"); // Disable GPU acceleration
			options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			options.addArguments("--start-maximized");	


       // System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
      

        int count = 0;
      
          String spValue = "";
          String finalSp = "";
          String offerValue = "NA";
          String newName = null;
          String mrpValue = null;
          String originalMrp1 = " ";
          String originalMrp2 = " ";
          String originalMrp3 = " ";
          String originalSp1 = " ";
          String originalSp2 = " ";
          String NewAvailability1 = " ";
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\Pharma Input1.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Apollo-3");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
	            		uRL = new ArrayList<>(),UOM = new ArrayList<>(),Mulitiplier = new ArrayList<>(),Availability = new ArrayList<>(),Pincode = new ArrayList<>(),NameForCheck = new ArrayList<>();
	            
            // Extract URLs from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                
                
                  if (i == 0) {
                continue;
            }    
                
                
                Cell inputPidCell = row.getCell(0);
                Cell inputCityCell = row.getCell(1);
                Cell inputNameCell = row.getCell(2);
                Cell inputSizeCell = row.getCell(3);
                Cell newProductCodeCell = row.getCell(4);
                Cell urlCell = row.getCell(5);
                Cell uomCell = row.getCell(6);
                Cell multiplierCell = row.getCell(7);
                Cell availabilityCell = row.getCell(8);
                Cell pinCodeCell = row.getCell(9);        
                Cell oldNameCell = row.getCell(10);    
                
            //   Cell urlCell = row.getCell(0);
              //  Cell urlCell = row.getCell(0);
               // Cell idCell = row.getCell(1);
               // Cell offerCell = row.getCell(2);
                
                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING) ? inputPidCell.getStringCellValue() : "";
                    String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING) ? inputCityCell.getStringCellValue() : "";
                    String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING) ? inputNameCell.getStringCellValue() : "";
                    String size = (inputSizeCell != null && inputSizeCell.getCellType() == CellType.STRING) ? inputSizeCell.getStringCellValue() : "";
                    String productCode = (newProductCodeCell != null && newProductCodeCell.getCellType() == CellType.STRING) ? newProductCodeCell.getStringCellValue() : "";
                    String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING) ? uomCell.getStringCellValue() : "";
                    String mulitiplier = (multiplierCell != null && multiplierCell.getCellType() == CellType.STRING) ? multiplierCell.getStringCellValue() : "";
                    String availability = (availabilityCell != null && availabilityCell.getCellType() == CellType.STRING) ? availabilityCell.getStringCellValue() : "";
                    String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";
                    String namecheck = (oldNameCell != null && oldNameCell.getCellType() == CellType.STRING) ? oldNameCell.getStringCellValue() : "";
                    
                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
                    UOM.add(uom);
                    Mulitiplier.add(mulitiplier);
                    Availability.add(availability);
                    Pincode.add(locationSet);
                    NameForCheck.add(namecheck);
                    
					/*
					 * uRL.add(url); ids.add(id); offers.add(offer);
					 */
                    
                }
            }
            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            Row headerRow = resultsSheet.createRow(0);
            
            
            headerRow.createCell(0).setCellValue("InputPid");
            headerRow.createCell(1).setCellValue("InputCity");
            headerRow.createCell(2).setCellValue("InputName");
            headerRow.createCell(3).setCellValue("InputSize");
            headerRow.createCell(4).setCellValue("NewProductCode");
            headerRow.createCell(5).setCellValue("URL");
            headerRow.createCell(6).setCellValue("Name");
            headerRow.createCell(7).setCellValue("MRP");
            headerRow.createCell(8).setCellValue("SP");
            headerRow.createCell(9).setCellValue("UOM");
            headerRow.createCell(10).setCellValue("Multiplier");
            headerRow.createCell(11).setCellValue("Availability");
            headerRow.createCell(12).setCellValue("Offer");
            headerRow.createCell(13).setCellValue("Commands");
            headerRow.createCell(14).setCellValue("Remarks");
            headerRow.createCell(15).setCellValue("Correctness");
            headerRow.createCell(16).setCellValue("Percentage");
            headerRow.createCell(17).setCellValue("Name");
            headerRow.createCell(18).setCellValue("Name Check");
            
            int rowIndex = 1;
            int headercount = 0;
            String currentPin = null;
            
            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String uom = UOM.get(i);
                String mulitiplier = Mulitiplier.get(i);
                String availability = Availability.get(i);
                String locationSet = Pincode.get(i);
                String namecheck = NameForCheck.get(i);
                
                
                
                try {
                	
                	  if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                          // Set "NA" values in all three columns
                          Row resultRow = resultsSheet.createRow(rowIndex++);
                          resultRow.createCell(0).setCellValue(id);
                          resultRow.createCell(1).setCellValue(city);
                          resultRow.createCell(2).setCellValue(name);
                          resultRow.createCell(3).setCellValue(size);
                          resultRow.createCell(4).setCellValue(productCode);
                          resultRow.createCell(5).setCellValue(url);
                          resultRow.createCell(6).setCellValue("NA");
                          resultRow.createCell(7).setCellValue("NA");
                          resultRow.createCell(8).setCellValue("NA");
                          resultRow.createCell(9).setCellValue("NA");
                          resultRow.createCell(10).setCellValue("NA");
                          resultRow.createCell(11).setCellValue("NA");
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  
                    driver.get(url);
                    driver.manage().window().maximize();
                    
                    if(i == 0) {
                    	Thread.sleep(10000);
                    }
                    
                    if (currentPin == null || !currentPin.equals(locationSet)) {
                     	try {
                     		
                     		//WebElement location1 = driver.findElement(By.className("LocationSearch_locationContainer__QNhiP"));
                     		
                     	Thread.sleep(3000);
                     	 WebElement location = driver.findElement(By.className("LocationSearch_locationContainer__QNhiP"));
      					location.click();
      					Thread.sleep(2000);     
      					WebElement loctionsetq = driver.findElement(By.xpath("//span[text()='Choose a different location']"));
      					loctionsetq.click();
//      					Thread.sleep(2000);
//      					 WebElement locationset = driver.findElement(By.xpath("/html/body/div[8]/div[2]/div/div[2]/div/div/div/div[1]/input"));
//       					locationset.click();
      					
      					
      					
                     }
                     	catch(Exception r) {
                     		Thread.sleep(3000);
                        	 WebElement location = driver.findElement(By.xpath("/html/body/div[8]/div[2]/div/div[2]/div/div/div/div[1]/input"));
         					location.click();
                     	}
      				//	Thread.sleep(1000);
      					
      					String tempPinNumber = "";
      					for (int j = 0; j < 200; j++) {
      						try {
      							driver.findElement(
      									By.xpath("//input[@placeholder='Search for society, locality, pincode...']"))
      									.sendKeys(Keys.ENTER);
      							
      							Thread.sleep(3000);
      							
      							driver.findElement(
      									By.xpath("//input[@placeholder='Search for society, locality, pincode...']")).clear();
       							Thread.sleep(3000);
      							
      							System.out.println("print the crt pin number" + locationSet);
      							
      							Thread.sleep(5000);
      							String crtPin = locationSet;
      							driver.findElement(
      									By.xpath("//input[@placeholder='Search for society, locality, pincode...']"))
      									.sendKeys(crtPin);
      							
      							
      							Thread.sleep(3000);
      							
      							currentPin = locationSet;
      							
      						/*	driver.findElement(
      									By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
      									.sendKeys(Keys.ENTER);   */
      							
      							for (int k = 0; k <= 50; k++) {
      								try {
      									tempPinNumber = driver.findElement(By.xpath(
      											"//div[@class='NewSearchLocationSuggestor_searchBlock__UPPqn']//div[@class='bL']"))
      											.getAttribute("value");
      									if (tempPinNumber.equals(locationSet)) {
      										break;
      									}
      								} catch (Exception e) {
      									if (i == 50) {
      										Assert.fail(e.getMessage());
      									}
      								}
      							}
      							
      							Thread.sleep(3000);
      							
      							driver.findElement(By.xpath("//div[@class='NewSearchLocationSuggestor_locationAutoComplete__PtF6Q']//ul//div[2]")).click();
      							
      							Thread.sleep(2000);
      							
      							currentPin = locationSet;
      							
      							break;
      						} catch (Exception e) {
      							e.getCause();
      							if (j == 300) {
      								Assert.fail(e.getMessage());
      							}
      						}
      					}
                          }
                     
                    try {
                    	
                    WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/h1"));
                    newName = nameElement.getText();
                    System.out.println(newName);
                    }
                    
                    catch(NoSuchElementException e) {
                    	
                    	WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/h1"));
                    	newName = nameElement.getText();
                        System.out.println(newName);
                    	
                    }
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                                      
                   try {
                    WebElement sp = driver.findElement(By.xpath("(//div[@id='PDP price banner']//div)[4]//p[1]"));
                    originalSp1 = sp.getText();
                    spValue =  originalSp1.replace("MRP: ₹", "").replace("*", "").replace("₹", "");
                
                    
                   }
                   catch(Exception e) {
                	   try {
                	   WebElement sp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  td']//p[1]"));
                       originalSp2 = sp.getText();
                       String spValue1 =  originalSp2.replace("*", "");
                       spValue = spValue1.replace("₹", "");
                       
                     //  System.out.println(spValue);
                	   }
                       catch(Exception exx) {
                    	   try {
                               WebElement sp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  ']//p[1]"));
                               originalSp1 = sp.getText();
                               spValue =  originalSp1.replace("MRP: ₹", "").replace("*", "").replace("₹", "");
                           //    System.out.println(spValue);
                               
                              }catch (Exception Ie) {
//                            	  WebElement sp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='ad hd id ']//p[1]"));
//                                  originalSp1 = sp.getText();
//                                  spValue =  originalSp1.replace("MRP: ₹", "").replace("*", "").replace("₹", "");
//                                  System.out.println(spValue);
                            	  
                            	  spValue="NA";
							}
                           }
                       }
                   System.out.println(spValue);
                   
                   try {
                       WebElement mrp = driver.findElement(By.xpath("((//div[@id='PDP price banner']//div)[4]//p[2]//span)[1]"));
                       originalMrp1 = mrp.getText();
                       mrpValue = originalMrp1.replace("MRP ₹", "");
                       } 
                       catch(NoSuchElementException e){ 
                       	try {
                       		
                       		WebElement mrp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  td']//p[2]"));
                               originalMrp2 = mrp.getText();
                               mrpValue = originalMrp2.replace("MRP ₹", "");
                        //       System.out.println(mrpValue);
                           
                       }
                       	catch(Exception ex) {
                       		try {
                                   WebElement mrp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  ']//p[2]"));
                                   originalMrp1 = mrp.getText();
                                   mrpValue = originalMrp1.replace("MRP ₹", "");
                         //          System.out.println(mrpValue);
                                   
                                   } catch (Exception FS) {
//                                   	 WebElement mrp = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='ad hd id ']//p[2]"));
//                                        originalMrp1 = mrp.getText();
//                                        mrpValue = originalMrp1.replace("MRP ₹", "");
//                                        System.out.println(mrpValue);
                                   	
                                   	mrpValue=spValue;
   								}
                       		}
                       	}

                   
                   System.out.println(mrpValue);

                   if(url.contains("NA")){
                	   String result = "NA";
                	   }
                	   String result="NA";
                	   try {
                	   //WebElement cartbtn = driver.findElement(By.xpath("//*[@id=\"add to cart banner\"]/div/div/div[2]/span"));
                	   WebElement cartbtn=driver.findElement(By.xpath("//*[contains(text(), 'This item is currently unavailable at your Pincode.')]"));
                	   if(cartbtn.isDisplayed()) {
                	   result="0";
                	   }}
                	   catch(Exception s){
                	   result="1";
                	   }
                	   //int stock = result;
                	   NewAvailability1 = String.valueOf(result);
                	   System.out.println(result);
                   // OFFER
                
					if(mrpValue.equals(spValue)){
	                	   offerValue = "NA";
	                   }
	                   else {
	                   try {
	                	   WebElement offer = driver.findElement(By.xpath("((//div[@id='PDP price banner']//div)[4]//p[2]//span)[2]"));
	                       String originalOffer = offer.getText();
	                       offerValue = originalOffer.replace("% off","% Off");
	                          System.out.println(offerValue);
	                      }
	                      catch(Exception e) {
	                    	  try {
	                    	  WebElement offer = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  ']//p[3]"));
	                          String getOffer = offer.getText();
	                          offerValue = getOffer.replace("% off","% Off");
	                          System.out.println(offerValue);
	                    	  }
	                    	  catch(Exception ex) {
	                    		  WebElement offer = driver.findElement(By.xpath("//div[@class='PdpWeb_addToCartSection__RGnBF']//div[@class='tL_  dd  td']//p[3]"));
	   	                       String originalOffer = offer.getText();
	   	                       offerValue = originalOffer.replace("% off","% Off");
	   	                          System.out.println(offerValue);
	                    	  }
	                      }    
	                   }
					
                   
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(name); 
                    resultRow.createCell(3).setCellValue(size);
                    resultRow.createCell(4).setCellValue(productCode);
                    resultRow.createCell(5).setCellValue(url);
                    resultRow.createCell(6).setCellValue(newName);
                    resultRow.createCell(7).setCellValue(mrpValue);
                    resultRow.createCell(8).setCellValue(spValue);
                    resultRow.createCell(9).setCellValue(uom);
                    resultRow.createCell(10).setCellValue(mulitiplier);
                    resultRow.createCell(11).setCellValue(NewAvailability1);
                    resultRow.createCell(12).setCellValue(offerValue);
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(" ");
                    resultRow.createCell(18).setCellValue(namecheck);
                   
                    
                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(name);
                    resultRow.createCell(3).setCellValue(size);
                    resultRow.createCell(4).setCellValue(productCode);
                    resultRow.createCell(5).setCellValue(url);
                    resultRow.createCell(6).setCellValue("NA");
                    resultRow.createCell(7).setCellValue("NA");
                    resultRow.createCell(8).setCellValue("NA");
                    resultRow.createCell(9).setCellValue(uom);
                    resultRow.createCell(10).setCellValue(mulitiplier);
                    resultRow.createCell(11).setCellValue(NewAvailability1);
                    resultRow.createCell(12).setCellValue(offerValue);
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(" ");
                    resultRow.createCell(18).setCellValue(namecheck);

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Apollo3_OutputData_" + timestamp + ".xlsx";
                
                // Write results to Excel file
                FileOutputStream outFile = new FileOutputStream(outputFilePath);
                resultsWorkbook.write(outFile);
                outFile.close();
                
                System.out.println("Output file saved: " + outputFilePath);
            } catch (Exception e) {
                e.printStackTrace();
            }
           
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
            	System.out.println("DoNe DoNe Scraping DoNe");
                driver.quit();
            }
        }
    }
}