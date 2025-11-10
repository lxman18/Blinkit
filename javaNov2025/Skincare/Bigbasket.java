package Daily_Run;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Bigbasket {

	public static void main(String[] args) throws Exception{
        //System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
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
          String currentPin = null;
          
        try {
            // Read URLs from Excel file
        	String filePath = ".\\input-data\\Website 8 Input data.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Bigbasket");//110015
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
                	  
                	  
                	 
  		    		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
  		    		WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));

  		    		
                		if (currentPin == null || !currentPin.equals(locationSet)) {
    						// Wait for and click the button to open the location setting
    						try {
    							// Navigate to BigBasket website
    							driver.get("https://www.bigbasket.com/");

    							// Wait for the location input field to be clickable (based on your HTML
    							// snippet)
    							WebDriverWait wait11 = new WebDriverWait(driver, Duration.ofSeconds(10));

    							try {

    							WebElement button = wait11.until(ExpectedConditions
    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[2]")));
    							button.click();
    							
    							}catch (Exception e) {
    								WebElement button = wait11.until(ExpectedConditions
	    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1]")));
	    							button.click();
								}
    							
    							try {
    								WebElement locationInput = wait.until(ExpectedConditions.elementToBeClickable(By
    										.xpath("//div[@class='flex flex-col absolute right-0 top-full mt-1.5 bg-white rounded-2xs outline-none z-max w-74 xl:w-90 scale-100']//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp']//input")));

    								// Click on the location input field
    								locationInput.click();
    								locationInput.sendKeys(locationSet);
    							} catch (Exception e) {

    								List<WebElement> searchInputElement = driver.findElements(
    										By.xpath("//input[@placeholder='Search for area or street name']"));
    								System.out.println("Search box size " + searchInputElement.size());
    								for (WebElement searchInput : searchInputElement) {
    									try {
    										searchInput.click();
    										searchInput.sendKeys(locationSet);
    									} catch (Exception e1) {
    										e1.printStackTrace();
    									}

    								}
//                                	 
    							}

    							// Wait for suggestions to appear
    							// Click the first XPath
    						
    							WebElement firstElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[1]")));
    							firstElement.click();

    							Thread.sleep(2000);
    							
    							// Wait for the message to appear
    							WebElement message = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(text(),'The selected city is not serviceable at the moment')]")));

    							if (message.isDisplayed()) {
    								
    								driver.get("https://www.bigbasket.com/");
    								try {

	    							WebElement button = wait1.until(ExpectedConditions
	    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[2]")));
	    							button.click();
	    							
	    							}catch (Exception e) {
	    								WebElement button = wait1.until(ExpectedConditions
		    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1]")));
		    							button.click();
									}
	    							
	    							try {
	    								WebElement locationInput = wait1.until(ExpectedConditions.elementToBeClickable(By
	    										.xpath("//div[@class='flex flex-col absolute right-0 top-full mt-1.5 bg-white rounded-2xs outline-none z-max w-74 xl:w-90 scale-100']//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp']//input")));

	    								// Click on the location input field
	    								locationInput.click();
	    								locationInput.sendKeys(locationSet);
	    							} catch (Exception e) {

	    								List<WebElement> searchInputElement = driver.findElements(
	    										By.xpath("//input[@placeholder='Search for area or street name']"));
	    								System.out.println("Search box size " + searchInputElement.size());
	    								for (WebElement searchInput : searchInputElement) {
	    									try {
	    										searchInput.click();
	    										searchInput.sendKeys(locationSet);
	    									} catch (Exception e1) {
	    										e1.printStackTrace();
	    									}

	    								}
	    
	    							}
	    						
    							    // If the message appears, click the second XPath
    							    WebElement secondElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[2]")));
    							    secondElement.click();
    							}

    						
    						
    							System.out.println("Location set to: " + locationSet);
    							
    							currentPin = locationSet;

    						} catch (Exception e) {
    							e.printStackTrace();
    							currentPin = locationSet;

    						}
    					}

                    driver.get(url);
                    driver.manage().window().maximize();
                    Thread.sleep(5000);
                    
                    try {
                    WebElement nameElement = driver.findElement(By.xpath("//h1[@class='Description___StyledH-sc-82a36a-2 bofYPK']"));
                    newName = nameElement.getText();
                    System.out.println(newName);
                    }
                    catch(NoSuchElementException e) {
                    	WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/section[1]/div[2]/section[1]/h1"));
                    	newName = nameElement.getText();
                        System.out.println(newName);
                    }
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    Thread.sleep(2000);
                    try {
                    WebElement mrp = driver.findElement(By.xpath("//td[@class='line-through p-0']"));
                    originalMrp1 = mrp.getText();
                    mrpValue = originalMrp1.replace("₹", "");
                    System.out.println(mrpValue);
                    }
                    
                    catch(NoSuchElementException e){ 
                    	try {
                    		
                    		WebElement mrp = driver.findElement(By.xpath("//td[@class='line-through p-0']"));
                            originalMrp2 = mrp.getText();
                            mrpValue = originalMrp2.replace("₹", "");
                           // mrpValue = originalMrp2;                      
                            System.out.println(mrpValue);
                        
                    }catch (Exception a) {
                    	try {
						WebElement bmrp = driver.findElement(By.xpath("//td[@class='Description___StyledTd-sc-82a36a-4 fLZywG']"));
						if(bmrp != null) {
						String MrpValue = bmrp.getText();
						Pattern pattern = Pattern.compile("₹(\\d+\\.?\\d*)");
	                    Matcher matcher = pattern.matcher(MrpValue);
	                    System.out.println(MrpValue);
	                    if (matcher.find()) {
	                        String extractedPrice = matcher.group(1);
	                       
	                        System.out.println(extractedPrice);
	                        mrpValue=extractedPrice;                    
	                   }}
                    	}catch(Exception ds) {
                    		mrpValue = "NA";
                    	}
					}
                    }
                    	
                    
                   try {
                    WebElement sp = driver.findElement(By.xpath("//td[@class='Description___StyledTd-sc-82a36a-4 fLZywG']"));
                    originalSp1 = sp.getText();
                    System.out.println("======================"+originalSp1+"====================");
                    Pattern pattern = Pattern.compile("₹(\\d+\\.?\\d*)");
                    Matcher matcher = pattern.matcher(originalSp1);
                    System.out.println(originalSp1);
                    
                    if (matcher.find()) {
                        String extractedPrice = matcher.group(1);
                       
                        System.out.println(extractedPrice);
                         spValue=extractedPrice;                    
                   }
                   }
                   catch(Exception e) {
                	   try {
                		   spValue = mrpValue;
                       System.out.println(spValue);
                	   }
                       catch(Exception exx) {
                    	   spValue = "NA";
                       }
                   }
                   
                   // offer
                   try {
                       WebElement offer = driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/div/div/section[1]/div[2]/section[1]/table/tr[3]/td[2]"));
                       String NewOffer = offer.getText();
                       
                   
                       
                       offerValue = NewOffer.replace("OFF", "Off");
                       System.out.println(offerValue);
                       
                      }
                      catch(Exception e) {
                    	  try {
                   	   WebElement offer = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]"));
                   	   String NewOffer = offer.getText();
                   	 offerValue = NewOffer.replace("OFF", "Off");
                          System.out.println(offerValue);
                    	  }
                    	  catch(Exception ex){
                    		  offerValue = "NA";
                    	  }
                      }    
                   
                   //Out Of Stocks
                   if(url.contains("NA")){
						String result = "NA";
					}	
			
                   int result=1;
                   
                   try {
                	   WebElement cart = driver.findElement(By.xpath("(//button[text()='Add to basket'])[1]"));
                	   result=1;
                   }catch (Exception e) {
                	   WebElement notify = driver.findElement(By.xpath("(//button[text()='Notify Me'])[1]"));
                	   result=0;
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
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
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
                String outputFilePath = ".\\Output\\Bigbasket_OutputData_" + timestamp + ".xlsx";
                
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
