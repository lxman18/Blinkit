package PharmaSunday;

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
import org.testng.Assert;


public class PharmaEasy_code_2ndhalf {

    public static void main(String[] args) {
    	 
		    	ChromeOptions options = new ChromeOptions();
		    	options.addArguments("--headless"); // Run Chrome in headless mode
		    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");	

		        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

		        // Schedule the task to run every day at 7:00 AM
		        Calendar now = Calendar.getInstance();
		        Calendar nextRunTime = Calendar.getInstance();
		        nextRunTime.set(Calendar.HOUR_OF_DAY, 15);
		        nextRunTime.set(Calendar.MINUTE, 30);
		        nextRunTime.set(Calendar.SECOND, 0);

		        long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		        if (initialDelay < 0) {
		            initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		        }

		        scheduler.scheduleAtFixedRate(() -> {
		            try {
		                System.out.println("Starting web scraping task...");
		                PharmaEasy_code_2ndhalf.runWebScraping();
		                System.out.println("Web scraping task completed.");
		            } catch (Exception e) {
		                e.printStackTrace();
		            }
		        }, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
		    }
		  		  

		    public static void runWebScraping() throws Exception{
		        	ChromeOptions options = new ChromeOptions();
			    	options.addArguments("--headless"); // Run Chrome in headless mode
			    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
			    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			    	options.addArguments("--start-maximized");
			    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			    	options.addArguments("--start-maximized");	

			    	//WebDriver driver = new ChromeDriver(options);

    	   WebDriver driver = new ChromeDriver();

	        int count = 0;
	        
	          String spValue = "";
	          String finalSp = "";
	          String offerValue = "NA";
	          String newName = null;
	          String newName1 = null;
	          String newName2 = null;
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
	            Sheet urlsSheet = urlsWorkbook.getSheet("PharmaEasy-2");//560010
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
	                        	Thread.sleep(8000);
	                        }
	                    if (i==0){
	                    	WebElement pincode=driver.findElement(By.xpath("//span[text()='Select Pincode']"));
	                    	pincode.click();
	                    	
	                    	  Thread.sleep(3000);
	                    	
	                    	WebElement inputbox=driver.findElement(By.xpath("//input[@placeholder='Enter PIN Code']"));
	                    	inputbox.click();
	                    	inputbox.sendKeys("560001");
	                    	
	                    	WebElement check=driver.findElement(By.xpath("//button[text()='Check']"));
	                    	check.click();
	                    }
	                    
	                    
	                     Thread.sleep(5000);
	                    try {
	                    	
	                    WebElement nameElement = driver.findElement(By.xpath("//h1"));
	                    newName = nameElement.getText();
	                    System.out.println(newName);
	                    
	                    }
	                    
	                    catch(NoSuchElementException e) {
	                    	
	                    	try {
								WebElement nameElement1 = driver.findElement(By.xpath("//h1[@class='MedicineOverviewSection_medicineName__9K61u']"));
								newName1 = nameElement1.getText();
								System.out.println(newName1);
							} catch (Exception ee) {
								WebElement nameElement2 = driver.findElement(By.xpath("//*[@id=\"__next\"]/main/div[1]/div[2]/div[2]/div[1]/div[2]/div/div/div/div[1]/h1"));
		                    	newName2 = nameElement2.getText();
		                        System.out.println(newName2);
		                    	
							}
	                    	
	                    	try {
								WebElement nameElement = driver.findElement(By.xpath("/html/body/div[1]/main/div[1]/div[2]/div[2]/div[1]/div[2]/div/div/div/div[1]/h1"));
								newName = nameElement.getText();
								System.out.println(newName);
							} catch (Exception e1) {
								System.out.println("NA");
							}
	                    	
	                    }
	                    System.out.println("headercount = " + headercount);
	                    
	                    headercount++;
	                    
	                    
	                    //sp
	                   try {
						try {
						    WebElement sp = driver.findElement(By.xpath("//span[@class='PriceInfo_unitPriceDecimal__i3Shz']"));
						    originalSp1 = sp.getText();
						    spValue =  originalSp1.replace("₹", "").replace(",", "");
						   }
						   catch(Exception e) {
							  // spValue = "NA";
							   try {
								   WebElement sp = driver.findElement(By.xpath("//div[@class='ProductPriceContainer_mrp___GtBV']"));
								   originalSp1 = sp.getText();
								   spValue =  originalSp1.replace("₹", "").replace(",", "");
							   }catch (Exception ef) {
								   WebElement sp = driver.findElement(By.xpath("//div[@class='PriceInfo_ourPrice__A549p']"));
								   originalSp1 = sp.getText();
								   spValue =  originalSp1.replace("₹", "").replace(",", "").replace("*", "");
							   	}
							   try {
								WebElement sp = driver.findElement(By.xpath("//*[@id=\"__next\"]/main/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/div/div[4]/div[1]/div/div/div[1]"));
								   originalSp1 = sp.getText();
								   spValue =  originalSp1.replace("₹", "").replace(",", "").replace("*", "");
							} catch (Exception e1) {
								// TODO Auto-generated catch block
								System.out.println("NA");
							}
							   
						       }
					} catch (Exception e) {
						System.out.println("NA");
					}
  	                    System.out.println(spValue);

	                   
	                   
	                   //mrp
	                   try {
	                       WebElement mrp = driver.findElement(By.xpath("//span[@class='PriceInfo_striked__fmcJv PriceInfo_costPrice__jhiax']"));
	                       originalMrp1 = mrp.getText();
	                       mrpValue = originalMrp1.replace("₹", "").replace(",", "").replace("M.R.P.:", "");
	                     
	                       } 
	                       
	                       catch(NoSuchElementException e){ 
	                    	   try {
	                    		   WebElement mrp = driver.findElement(By.xpath("//span[@class='ProductPriceContainer_striked__jDAwy']"));
	    	                       originalMrp1 = mrp.getText();
	    	                       mrpValue = originalMrp1.replace("₹", "").replace(",", "").replace("M.R.P.:", "");
	                    	   }catch (Exception es) {
	                    		   mrpValue =spValue;
							}
	                       	}
	                   
                       System.out.println(mrpValue);
	                   
	                 //Out Of Stocks
	                   int result=1;
	               try {
	                	   WebElement cart=driver.findElement(By.xpath("//button[text()='Add To Cart']"));
	                	   
	                   if(cart.isDisplayed() && cart.isEnabled()) {
	                	   result=1;
	                   }}
	               catch (Exception e) {
	            	   try {
						WebElement cart=driver.findElement(By.xpath("(//button[@id='proceed']//span)[1]"));
						   result=1;
					} catch (Exception e1) {
						WebElement cart=driver.findElement(By.xpath("//button[text()='Notify Me']"));
						   result=0;
					}

				}
	                   
					NewAvailability1 = String.valueOf(result);
					System.out.println(NewAvailability1);

	                   
						//offer
						if(mrpValue.equals(spValue))
						{
							offerValue="NA";
						}else {
						try {
							WebElement offer=driver.findElement(By.xpath("//div[@class='PriceInfo_gcdDiscountPercent__FvJsG']"));
							offerValue=offer.getText();
						}catch (Exception e) {
							WebElement offer=driver.findElement(By.xpath("//DIV[@class='ProductPriceContainer_discountContainer__5arLQ']"));
							offerValue=offer.getText();					
							}}
						System.out.println(offerValue);
	                 
	                   
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
	                String outputFilePath = ".\\Output\\pharma_Easy2_Output_" + timestamp + ".xlsx";
	                
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

