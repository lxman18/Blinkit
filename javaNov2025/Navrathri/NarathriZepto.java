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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.util.Calendar;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

public class NarathriZepto {
	
	public static void main(String[] args) {
			    	//WebDriver driver = new ChromeDriver(options);

	       // System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
	        WebDriver driver = new ChromeDriver();


	        int count = 0;
	        // int finalSp;
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
	            String filePath = ".\\input-data\\Navrathri-Input.xlsx";
	            FileInputStream file = new FileInputStream(filePath);
	            Workbook urlsWorkbook = new XSSFWorkbook(file);
	            Sheet urlsSheet = urlsWorkbook.getSheet("Zepto");
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
	                
	            //    Cell urlCell = row.getCell(0);
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
	            headerRow.createCell(12).setCellValue("Commands");
	            headerRow.createCell(13).setCellValue("Remarks");
	            headerRow.createCell(14).setCellValue("Correctness");
	            headerRow.createCell(15).setCellValue("Percentage");
	            headerRow.createCell(16).setCellValue("Name");
	            headerRow.createCell(17).setCellValue("Offer");
	            headerRow.createCell(18).setCellValue("NameForCheck");
	            
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
	                          resultRow.createCell(12).setCellValue("NA");
	                          
	                          System.out.println("Skipped processing for URL: " + url);
	                          continue; // Skip to the next iteration
	                      }
	                	  
	                    driver.get(url);
	                    Thread.sleep(5000);
	                    driver.manage().window().maximize();
	                   
	        
	                    //location sets 
	         /*           if (currentPin == null || !currentPin.equals(locationSet)) {
	                        
	                        Thread.sleep(9000); 	
	                          	
	                          	
	                         for(int k=0;k < 100;k++) {
	                          	try {
	                          	Thread.sleep(3000);
	                          	 WebElement location = driver.findElement(By.xpath("//button[@aria-label='Select Location']"));
	           					location.click();
	           					break;
	                          }
	                          	catch(Exception r) {
	                          			try {
	                          		Thread.sleep(4000);
	                             	 WebElement location = driver.findElement(By.xpath("/html/body/div[1]/div/header/div/div[2]/button"));          			 
	              					location.click();
	              					break;
	              					
	                          			}
	                          			catch(NoSuchElementException u) {
	                          				try {
	                          				Thread.sleep(4000);
	                                        	 WebElement location = driver.findElement(By.className("/html/body/div[2]/div/div/div/div[2]/div/button[1]/div/p"));          			 
	                         					location.click();
	                         					Thread.sleep(3000);
	                         					break;
	                          			}
	                          				catch(Exception tr) {
	                          					Thread.sleep(4000);
	                                           	 WebElement location = driver.findElement(By.xpath("/html/body/div[1]/div/header/div/div[1]/button"));          			 
	                            					location.click();
	                            					Thread.sleep(3000);
	                            					break;
	                          				}
	                          				//div[@class = 'flex flex-col justify-center relative lg:flex-initial flex-1 h-[2.75rem]']//following::span[2]
	                          			}
	                          	}
	                         }
	                          	
	           				//	Thread.sleep(1000);909
	                          	
	           					String tempPinNumber = "";
	           					for (int j = 0; j < 20; j++) {
	           						try {         
	           							
	           							
	           							
	           							Thread.sleep(3000);
	           							
	           							driver.findElement(
	           									By.xpath("//input[@placeholder='Search a new address']"))
	           									.sendKeys(Keys.ENTER);
	           							
	           							Thread.sleep(3000);
	           							
	           							driver.findElement(
	           									By.xpath("//input[@placeholder='Search a new address']")).clear();
	           							
	           							Thread.sleep(3000);
	           							
	           							System.out.println("print the crt pin number" + locationSet);
	           							
	           							String crtPin = locationSet;
	           							driver.findElement(
	           									By.xpath("//input[@placeholder='Search a new address']"))
	           									.sendKeys("560001");
	           							
	           							Thread.sleep(3000);
	           							
	           							try {
	           							
	           							driver.findElement(By.xpath("/html/body/div[4]/div/div/div/div/div/div/div[2]/div/div[2]/div[2]/div[1]")).click();
	           							}
	           							catch (Exception e) {
		           							driver.findElement(By.xpath("(//div[@data-testid='address-search-item'])[1]")).click();
										}
	           							Thread.sleep(2000);
	           							
	           							currentPin = locationSet;
	           							
	           							
	           							System.out.println("============="  + currentPin+ "======================");
	           							
	           						/*	driver.findElement(
	           									By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
	           									.sendKeys(Keys.ENTER);   
	           							
	           							Thread.sleep(3000);//
           							
           							driver.findElement(By.xpath("//div[contains(text(), 'Confirm & Continue')]")).click();
           							
	           							Thread.sleep(2000);
           							
           							currentPin = locationSet;
           							
           							
           							
           							break;
           						} 
           						
						catch (Exception e) {
           				
           						}
           						}
           				
           											}*/
	                    
	                    
	                  if(i==0) {
	                    	try {
	                    		 WebElement location = driver.findElement(By.xpath("//button[@aria-label='Select Location']"));
		           					location.click();
	                    	}catch (Exception e) {
	                    		 WebElement location = driver.findElement(By.xpath("/html/body/div[1]/div/div/header/div/div[2]/button/span/span"));
	                    		 location.click();
							}
	                    	
	                    	try
	                    	{
	                    		Thread.sleep(2000);
	                    		
	                    		driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']"))
       									.sendKeys(Keys.ENTER);
	                    		driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']")).clear();
       							
       							Thread.sleep(3000);
       							
       							System.out.println("print the crt pin number" + locationSet);
       							
       						
       							driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']"))
       									.sendKeys("110015");
	                    	}
	                    	catch (Exception e) {

	                    		driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']"))
       									.sendKeys(Keys.ENTER);
	                    		driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']")).clear();
       							
       							Thread.sleep(3000);
       							
       							System.out.println("print the crt pin number" + locationSet);
       							
       						
       							driver.findElement(
       									By.xpath("//input[@placeholder='Search a new address']"))
       									.sendKeys("110015");
							}
	                    	
	                    	Thread.sleep(3000);
	                    	
	                    	try {
       							
       							driver.findElement(By.xpath("/html/body/div[4]/div/div/div/div/div/div/div[2]/div/div[2]/div[2]/div[1]")).click();
       							}
       							catch (Exception e) {
           							driver.findElement(By.xpath("(//div[@data-testid='address-search-item'])[1]")).click();
								}
	                    	Thread.sleep(3000);

   							driver.findElement(By.xpath("//button[contains(text(), 'Confirm & Continue')]")).click();

	                    }
	
	               	                    
	                    Thread.sleep(2000);
	                       
	                    try {
	                    	
	                    WebElement nameElement = driver.findElement(By.xpath("//div[@class='mt-2 flex items-center justify-between gap-6']//h1"));
	                    newName = nameElement.getText();
	                    System.out.println(newName);
	                    }
	                    
	                    catch(NoSuchElementException e) {
	                    	
	                    	WebElement nameElement = driver.findElement(By.xpath("/html/body/div/main/div[2]/div/div[2]/div[2]/div/h1"));
	                    	newName = nameElement.getText();
	                        System.out.println(newName);
	                    	
	                    }
	                    System.out.println("headercount = " + headercount);
	                    
	                    headercount++;

	           //sp
	                    Thread.sleep(2000);
	                   try {
	                    WebElement sp = driver.findElement(By.xpath("//span[@class='text-[32px] font-medium leading-[30px] text-[#262A33]']"));
	                    originalSp1 = sp.getText();
	                    spValue =  originalSp1.replace("₹", "");
	                    System.out.println(spValue);
	                   }
	                   catch(Exception e) {
	                	   
	                	   try {
	                	   WebElement sp = driver.findElement(By.xpath("//span[@class='text-[32px] font-medium leading-[30px] text-[#262A33]']"));
	                       originalSp2 = sp.getText();
	                       spValue =  originalSp2.replace("₹", "");
	                       System.out.println(spValue);
	                	   }
	                       catch(Exception exx) {
	                    	   spValue = "NA";
	                       }
	                   }//(//div[@class='flex items-center gap-2'])[2]//span
	                   
	                   //mrp
	                   Thread.sleep(2000);
	                    try {
	                    WebElement mrp = driver.findElement(By.xpath("//span[@class='line-through font-bold']"));
	                    originalMrp1 = mrp.getText();
	                    mrpValue = originalMrp1.replace("₹", "");
	                    System.out.println(mrpValue);
	                    
	                    }
	                    
	                    catch(NoSuchElementException e){ 
	                    	try {
	                    		
	                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\"product-features-wrapper\"]/div[1]/div/div[3]/div[1]/div[2]/p/span[2]"));
	                            originalMrp2 = mrp.getText();
	                            mrpValue = originalMrp2.replace("₹", "");
	                           // mrpValue = originalMrp2;                      
	                            System.out.println(mrpValue);
	                        
	                    }
	                    	catch(Exception ex) {
	                    		try {
	                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
	                           // WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div/div[7]/div[3]/div[4]/div[12]/div/div/div[4]/div[2]/span/span[1]/span[2]/span/span[2]"));
	                            originalMrp3 = mrp.getText();
	                            mrpValue = originalMrp3.replace("₹", "");                      
	                            System.out.println(mrpValue);
	                    		}
	                    		catch(Exception exx) {
	                    			mrpValue = spValue;
	                    			System.out.println(mrpValue);
	                    		}
	                    		}
	                    	}
	                    
	                   Thread.sleep(2000);
   
						//offer
	                   
//	                   if(mrpValue.equals(spValue)) {
//	                	   offerValue="NA";
//	                   }
//	                   System.out.println(offerValue);
	                   
	                   try {
	                    	
		                    WebElement offer = driver.findElement(By.xpath("//div[@class='flex items-center gap-2']//p[2]"));
		                    offerValue = offer.getText();
		                    }
		                    
		                    catch(NoSuchElementException e) {
		                    	try {
		                    	WebElement offer = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div[2]/div[3]/div[2]/div[1]/div[1]/div/div[4]/div[1]/div[1]/div//p[2]"));
		                    	offerValue = offer.getText();
		                       
		                    	}
		                    	catch (Exception se) {
		                    		offerValue="NA";
								}
		                    }
	                   System.out.println(offerValue);
	                   
						
						if(url.contains("NA")){
							String result = "NA";
						}	
						
						int result = 1;

											
						try {
							WebElement cart=driver.findElement(By.xpath("(//button[contains(text(), 'Add To Cart')])[2]"));
							if(cart.isEnabled()) {
								result=1;
							}
							else {
								result=0;
							}
							
						}catch (Exception ae) {
							WebElement notify=driver.findElement(By.xpath("(//span[text()='Notify Me'])[2]"));

								if(notify.isDisplayed()) {
									result=0;
								}
						}
						   System.out.println(result);
						NewAvailability1 = String.valueOf(result);

	                   
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
	            
	            // Write results to Excel file
	            String timestamp = new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss").format(new Date());
	            FileOutputStream outFile = new FileOutputStream(".\\Output\\Zepto_Navrathri_Output"+timestamp+".xlsx");
	            resultsWorkbook.write(outFile);
	            outFile.close();

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