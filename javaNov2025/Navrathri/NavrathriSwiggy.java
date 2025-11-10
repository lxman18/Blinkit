import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
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

public class NavrathriSwiggy {

    public static void main(String[] args) throws Exception{
      
        WebDriver driver = new ChromeDriver();

        int count = 0;
        // int finalSp;
          String spValue = "";
          String finalSp = "";
          String offerValue = "NA";
          String newName = null;
          String brandName=null;
          String pname=null;
          String mrpValue = null;
          String originalMrp1 = " ";
          String originalMrp2 = " ";
          String originalMrp3 = " ";
          String originalSp1 = " ";
          String originalSp2 = " ";
          String uomNew = " ";
          String NewAvailability1 = " ";
          String off = ""; 
          
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\Navrathri-Input.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Swiggy1");
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
            String currentPin =null;
            
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
                	  
                    //location sets 
                	  
           	  if (currentPin == null || !currentPin.equals(locationSet)) {
                		  
                		  
                		  driver.get("https://www.swiggy.com/");
                		  driver.manage().window().maximize();
                          
                          //location sets 
                		//  WebElement location = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/header/div/div/div/span[1]"));
                		  
                		  WebElement location = driver.findElement(By.xpath("/html/body/div/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input"));
        					location.click();
        					String tempPinNumber = "";
        					for (int j = 0; j < 150; j++) {
        						try {
        							driver.findElement(
        									By.xpath("//*[@id=\"location\"]"))
        									.sendKeys(Keys.ENTER);
        							
        							Thread.sleep(1000);
        							
        							WebElement locationField = driver.findElement(By.xpath("//*[@id=\"location\"]"));

        							// Create an Actions object
        							Actions actions = new Actions(driver);

        							// Perform Ctrl + A (Select All) and then clear the field
        							actions.moveToElement(locationField) // Move to the input field
        							       .click()                    // Focus on the input field
        							       .keyDown(Keys.CONTROL)      // Press Ctrl
        							       .sendKeys("a")              // Press A (to select all)
        							       .keyUp(Keys.CONTROL)        // Release Ctrl
        							       .perform();     
        							Thread.sleep(1500);
        							
        							
        							locationField.sendKeys(Keys.DELETE);
       							
        							
        							Thread.sleep(1000);
        							
        							System.out.println("print the crt pin number" + locationSet);
        							
        							String crtPin = locationSet;
        							driver.findElement(
        									By.xpath("//*[@id=\"location\"]"))
        									.sendKeys(crtPin);
        							
        							
        							Thread.sleep(1000);
        							
        							currentPin = locationSet;
        							
        						/*	driver.findElement(
        									By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
        									.sendKeys(Keys.ENTER);   
        							
        							for (int k = 0; k <= 50; k++) {
        								try {
        									tempPinNumber = driver.findElement(By.xpath(
        											"//*[@id=\"location\"]"))
        											.getAttribute("value");
        									if (tempPinNumber.equals(locationSet)) {
        										break;
        									}
        								} catch (Exception e) {
        									if (i == 50) {
        										Assert.fail(e.getMessage());
        									}
        								}
        							}*/
        							
        							Thread.sleep(1000);
        							
        							driver.findElement(By.xpath("(//div[@class='_3bVtO']//div[@class='kuQWc'])[1]")).click();
        							
        							Thread.sleep(1000);
        							
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
                    
                	  
                
            		  
               /* 	  String crtPin = locationSet;
                	  
                	  if (currentPin == null || !currentPin.equals(locationSet)) {
                		  
                		  driver.get("https://www.swiggy.com/");
                		  driver.manage().window().maximize();
                		  
                		  WebElement pin= driver.findElement(By.xpath("//span[text()='Other']"));
                		  pin.click();
                			Thread.sleep(1000);
                		  
                		  WebElement locate=driver.findElement(By.xpath("//input[@placeholder='Search for area, street name..']"));
                		  locate.click();
                		  locate.sendKeys(crtPin);
                			Thread.sleep(2000);
                		  
                		  currentPin = locationSet;
                		  
                		  WebElement drop=driver.findElement(By.xpath("//div[@class='_2RwM6']"));
                		  drop.click();
                		  
                	  }*/
                      
                	  driver.get(url);
                   
                	  
                	Thread.sleep(1000);  
                	
                	 try {
                         WebElement wrong = driver.findElement(By.xpath("//div[text()='Something went wrong!']"));
                         if (wrong.isDisplayed()) {
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
                             resultRow.createCell(13).setCellValue(" ");
                             resultRow.createCell(14).setCellValue(" ");
                             resultRow.createCell(15).setCellValue(" ");
                             resultRow.createCell(16).setCellValue(" ");
                             resultRow.createCell(17).setCellValue(offerValue);
                             resultRow.createCell(18).setCellValue(namecheck);

                             System.out.println("Something went wrong found, skipping URL: " + url);
                             continue; // Skip further processing for this URL
                         }
                	 	} catch (NoSuchElementException ew) {
                	  try {
                          WebElement nameElement = driver.findElement(By.xpath("//div[@data-testid='item-name']"));
                          newName = nameElement.getText();
                          System.out.println(newName);
                          }
                          catch(NoSuchElementException e) {
                          	WebElement nameElement = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/div[2]/div[1]/div/div/div[2]/div[2]/div[2]"));
                          	newName = nameElement.getText();
                              System.out.println(newName);
                          }
                          System.out.println("headercount = " + headercount);
                          
                          headercount++;
                          
                          try {
                              WebElement sp = driver.findElement(By.xpath("(//div[@class='_AHZN Hxf8j']//div//div)[2]//div[@data-testid='item-offer-price']"));
                              originalSp1 = sp.getText();
                              spValue =  originalSp1.replace("₹", "");
                              System.out.println(spValue);
                             }
                             catch(Exception e) {
                            	 WebElement sp = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/div[2]/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]"));
                                 originalSp1 = sp.getText();
                                 spValue =  originalSp1.replace("₹", "");
                                 System.out.println(spValue);
                             }
                          // Mrp 
                          try {
                              WebElement mrp = driver.findElement(By.xpath("(//div[@class='_AHZN Hxf8j']//div//div)[2]//div[@data-testid='item-mrp-price']"));
                              originalMrp1 = mrp.getText();
                              mrpValue = originalMrp1.replace("₹", "");
                              System.out.println(mrpValue);
                              
                              }
                              
                              catch(NoSuchElementException e){ 
                           	   try {
                                      WebElement mrp = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/div[2]/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]"));
                                      originalMrp1 = mrp.getText();
                                      mrpValue = originalMrp1.replace("₹", "");
                                      System.out.println(mrpValue);
                                      
                                      }
                           	   catch (Exception S) {
                           		mrpValue=spValue;
                           		System.out.println(mrpValue);
       						}
                                      
                          	}
                          
                          //Offer 
                          try {
                              WebElement Offer = driver.findElement(By.xpath("(//div[@data-testid='item-offer-label-discount-text'])[1]"));
                              offerValue = Offer.getText();
                             
                              System.out.println(offerValue);
                              
                              }
                              
                              catch(NoSuchElementException e){ 
                           	   try {
                           		 WebElement Offer = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div"));
                                 offerValue = Offer.getText();
                                
                                 System.out.println(offerValue);
                                      
                                      }
                           	   catch (Exception S) {
                           		mrpValue=spValue;
                           		System.out.println(mrpValue);
       						}
                                      
                          	}
                   //out Of Stock
                   
                   int result=1;
                   if (url.contains("NA")) {
                	   NewAvailability1 = "NA";
                	   } 
                   else {
                	 
                	   try {
                	   // Define the texts to check for
                		   String[] textsToCheck = {
                				   "Currently Unavailable",
                				   "Currently out of stock in this area.",
                				   "Sold Out",
                				   "Unavailable"
                				   };

                	   // Get the page source
                	   String pageSource = driver.getPageSource();
                	   boolean isTextPresent = false;

                	   // Check for the presence of any of the texts
                	   for (String text : textsToCheck) {
                	   if (pageSource.contains(text)) {
                	   isTextPresent = true;
                	   break;
                	   }
                	   }

                	   // Determine the result based on the presence of the text
                	   result = isTextPresent ? 0 : 1;
                	   System.out.println(result);
                	   } catch (Exception e) {
                	   System.out.println("Error checking availability: " + e.getMessage());
                	   result = -1;
                	   }
                	   }

                	   // Assign final availability status
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
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
                    resultRow.createCell(18).setCellValue(namecheck);
                    
                    System.out.println("Data extracted for URL: " + url);
                     }   } catch (Exception e) {
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
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
                    resultRow.createCell(18).setCellValue(namecheck);
                    

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Swiggy1_Narathri_Output_" + timestamp + ".xlsx";
                
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


