package scheduleRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

		public class HeadsUp_fortail2_Bang{
		    public static void main(String[] args) {
		    	ChromeOptions options = new ChromeOptions();
		    	//options.addArguments("--headless"); // Run Chrome in headless mode
		    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");

		        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

		        // Schedule the task to run every day at 7:00 AM
		        Calendar now = Calendar.getInstance();
		        Calendar nextRunTime = Calendar.getInstance();
		        nextRunTime.set(Calendar.HOUR_OF_DAY, 3);
		        nextRunTime.set(Calendar.MINUTE, 15);
		        nextRunTime.set(Calendar.SECOND, 0);

		        long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		        if (initialDelay < 0) {
		            initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		        }

		        scheduler.scheduleAtFixedRate(() -> {
		            try {
		                System.out.println("Starting web scraping task...");
		                HeadsUp_fortail2_Bang.runWebScraping();
		                System.out.println("Web scraping task completed.");
		            } catch (Exception e) {
		                e.printStackTrace();
		            }
		        }, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
		    }


		    public static void runWebScraping() throws Exception{
		        	ChromeOptions options = new ChromeOptions();
		      //  	options.addArguments("--headless"); // Run Chrome in headless mode
			    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
			    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			    	options.addArguments("--start-maximized");
			    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
			    	options.addArguments("--start-maximized");

			    	//WebDriver driver = new ChromeDriver(options);

			       // System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

			        WebDriver driver = new ChromeDriver(options);
			    // 	EdgeOptions option = new EdgeOptions();
			     //	WebDriver driver = new EdgeDriver(option);
			         driver.manage().window().maximize();
			         WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

			         List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
			                 InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
			                 UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(),
			                 Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();

			         Workbook resultsWorkbook = new XSSFWorkbook();
			         Sheet resultsSheet = resultsWorkbook.createSheet("Results");
			         createHeaderRow(resultsSheet);

			         int rowIndex = 1;
			         String currentPin = null;

			         try (FileInputStream file = new FileInputStream(".\\input-data\\headsupforTail Input data.xlsx");
			              Workbook urlsWorkbook = new XSSFWorkbook(file)) {

			             Sheet urlsSheet = urlsWorkbook.getSheet("HeadsUp2");//Bang
			             int rowCount = urlsSheet.getPhysicalNumberOfRows();

			             // Extract URLs from Excel
			             for (int i = 1; i < rowCount; i++) {
			                 Row row = urlsSheet.getRow(i);
			                 if (row.getCell(5) != null) {
			                     String url;
			                     if (row.getCell(5).getCellType() == CellType.STRING) {
			                         url = row.getCell(5).getStringCellValue();
			                     } else if (row.getCell(5).getCellType() == CellType.NUMERIC) {
			                         url = String.valueOf(row.getCell(5).getNumericCellValue());
			                     } else {
			                         url = "NA"; // Handle cases where the cell is neither numeric nor string
			                     }

			                     inputPid.add(row.getCell(0).getStringCellValue());
			                     InputCity.add(row.getCell(1).getStringCellValue());
			                     InputName.add(row.getCell(2).getStringCellValue());
			                     InputSize.add(row.getCell(3).getStringCellValue());
			                     NewProductCode.add(row.getCell(4).getStringCellValue());
			                     uRL.add(url);
			                     UOM.add(row.getCell(6).getStringCellValue());
			                     Mulitiplier.add(row.getCell(7).getStringCellValue());
			                     Pincode.add(row.getCell(9).getStringCellValue());
			                     NameForCheck.add(row.getCell(10).getStringCellValue());
			                 }
			             }

			             int ProductCOUNT = 0;



			             // Main data extraction logic
			             for (int i = 0; i < uRL.size(); i++) {
			                 String url = uRL.get(i);
			                 String locationSet = Pincode.get(i); // Get the current pincode from the list
			                 if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
			                     writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
			                             InputSize.get(i), "NA", url, "NA", "NA", "NA",
			                             "NA", Mulitiplier.get(i), "NA", "NA", "NA", "NA", "NA", "NA", NameForCheck.get(i));
			                     System.out.println("Skipped processing for URL: " + url);
			                     continue;
			                 }

			                 try {
			                     driver.get(url);
			                     Thread.sleep(2000);

			                     if (currentPin == null || !currentPin.equals(locationSet)) {
			                         Thread.sleep(2000);
			                         driver.findElement(By.xpath("//div[@class='dtl_imputs']//input")).click();
			                         Thread.sleep(2000);
			                         driver.findElement(By.xpath("//div[@class='dtl_imputs']//input")).click();
			                         Thread.sleep(2000);
			                         driver.findElement(By.xpath("//div[@class='dtl_imputs']//input")).sendKeys(locationSet);
			                         Thread.sleep(2000);
			                         driver.findElement(By.xpath("//button[.='Check']")).click();
			                         Thread.sleep(2000);

			                         currentPin = locationSet;

			                         driver.get(url);
			                     }
			                     Thread.sleep(3000);

			                  // Check if "Page not found" exists

			                     String pageTitle = driver.getTitle();

			                     // Print the page title
			                     System.out.println("Page title is: " + pageTitle);

			                     if (pageTitle.startsWith("404 Not Found ") || pageTitle.startsWith("Heads Up For Tails")) {
			                    	    writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
			                    	            InputSize.get(i), "NA", "NA", "NA", "NA", "NA",
			                    	            "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", NameForCheck.get(i));
			                    	    System.out.println("Page title indicates no data to extract for URL: " + url);
			                    	    continue;  // Skip further processing for this URL

			                     } else {
			                         // If not found, proceed as usual



			                     // Extract product ID
			                     String productId = extractProductId(url);
			                     System.out.println("Extracted Product ID: " + url);

			                     String newName = " ";
			                     try {
			                         newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='product__title']//h1"))).getText();
			                         System.out.println(newName);
			                     } catch (Exception Q) {
			                         newName = "NA";
			                     }

			                     // Extract MRP
			                     String mrpValue = extractMRP(driver);
			                     System.out.println(mrpValue);

			                     // Extract SP
			                     String spValue = extractSP(driver);
			                     System.out.println(spValue);

			                     if ("NA".equals(spValue) && !mrpValue.equals("NA")) {
			                         spValue = mrpValue; // Assign MRP to SP if SP is NA and MRP is valid
			                     }


			                     // Extract offer
			                     String offerValue = extractOffer(driver);
			                     System.out.println(offerValue);

			                     // Take screenshot using productId from inputPid list
			                     takeScreenshot(driver, inputPid.get(i));  // Pass the productId from inputPid list

			                     int result=1;

			                     try {
			                     	WebElement cart=driver.findElement(By.xpath("(//div[@class='product-form__buttons']//button)[1]"));
			                     	if(cart.isDisplayed() && cart.isEnabled()) {
			                     		result=1;
			                     	}
			                     	else {
			                     		result=0;
			                     	}
			                     }catch (Exception eEE) {
			                     	WebElement cart=driver.findElement(By.xpath("(//div[@class='product-form__buttons']//button)[2]"));
			                     	if(cart.isDisplayed()) {
			                     		result=0;
			                     	}
			 					}

			                     String availability = String.valueOf(result);
			                     System.out.println(availability);

			                     ProductCOUNT++;

			                     // Write results to the results sheet
			                     writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
			                             InputSize.get(i), productId, url, newName, mrpValue, spValue,
			                             extractWeight(UOM, i), String.valueOf(Mulitiplier.get(i)), availability, offerValue, "NA", "NA", "NA", "NA", NameForCheck.get(i));

			                     System.out.println("Data extracted for URL: " + url);
			                 }   }catch (Exception e) {
			                     e.printStackTrace();
			                     System.out.println("Failed to extract data for URL: " + url);
			                 }
			             }
			         } catch (Exception e) {
			             e.printStackTrace();
			             System.out.println("An error occurred during the extraction process.");
			         } finally {
			             // Write results to Excel file
			             saveResultsToExcel(resultsWorkbook);

			             // Clean up driver resources
			             if (driver != null) {
			                 System.out.println("Closing the driver.");
			                 driver.quit();
			             }
			         }
			     }

			     private static void createHeaderRow(Sheet sheet) {
			         Row headerRow = sheet.createRow(0);
			         String[] headers = {"InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP",
			                 "SP", "UOM", "Multiplier", "Availability", "Offer", "Commands", "Remarks", "Correctness",
			                 "Percentage", "Name", "NameForCheck"};
			         for (int i = 0; i < headers.length; i++) {
			             headerRow.createCell(i).setCellValue(headers[i]);
			         }
			     }

			     private static void writeResults(Sheet sheet, int rowIndex, String id, String city, String name, String size,
			                                       String productId, String url, String newName, String mrpValue,
			                                       String spValue, String weight, String mulitiplier, String availability, String offerValue,
			                                       String command, String remark, String correctness, String percentage,
			                                       String nameCheck) {
			         Row resultRow = sheet.createRow(rowIndex);
			         resultRow.createCell(0).setCellValue(id);
			         resultRow.createCell(1).setCellValue(city);
			         resultRow.createCell(2).setCellValue(name);
			         resultRow.createCell(3).setCellValue(size);
			         resultRow.createCell(4).setCellValue(productId);
			         resultRow.createCell(5).setCellValue(url);
			         resultRow.createCell(6).setCellValue(newName);
			         resultRow.createCell(7).setCellValue(mrpValue);
			         resultRow.createCell(8).setCellValue(spValue);
			         resultRow.createCell(9).setCellValue(weight);
			         resultRow.createCell(10).setCellValue(mulitiplier);
			         resultRow.createCell(11).setCellValue(availability);
			         resultRow.createCell(12).setCellValue(offerValue);
			         resultRow.createCell(13).setCellValue(command);
			         resultRow.createCell(14).setCellValue(remark);
			         resultRow.createCell(15).setCellValue(correctness);
			         resultRow.createCell(16).setCellValue(percentage);
			         resultRow.createCell(17).setCellValue(nameCheck);
			     }

			     private static void saveResultsToExcel(Workbook resultsWorkbook) {
			         try {
			             SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
			             String timestamp = dateFormat.format(new Date());
			             String outputFilePath = ".\\Output\\HeadUpTail_OutputData_Bang" + timestamp + ".xlsx";

			             // Write results to Excel file
			             try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
			                 resultsWorkbook.write(outFile);
			             }

			             System.out.println("Output file saved: " + outputFilePath);
			         } catch (Exception e) {
			             e.printStackTrace();
			             System.out.println("Failed to save the output file.");
			         }
			     }

			     private static String extractMRP(WebDriver driver) {
			         try {
			             // Try to extract MRP from the price__sale section first
			             WebElement element = null;

			             try {
			                 // Look for the MRP in the price__sale section using the updated XPath
			                 element = driver.findElement(By.xpath("//div[@class='price__sale']//span//s[@class='price-item price-item--regular']"));
			             } catch (NoSuchElementException e) {
			                 // If the element is not found in this section, set it to null
			                 element = null;
			             }

			             // If an element was found in the price__sale section, check if it's a valid value
			             if (element != null && !element.getText().isEmpty()) {
			                 String text = element.getText();
			                 return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
			             }

			             // If the first XPath fails or has no valid MRP, now check the price__regular section
			             try {
			                 // Look for the MRP in the price__regular section
			                 element = driver.findElement(By.xpath("//div[@class='price__regular']//span[@class='price-item price-item--regular']"));
			             } catch (NoSuchElementException e) {
			                 // If the element is not found in this section either, return "NA"
			                 element = null;
			             }

			             // If the element is found in the price__regular section, return the MRP
			             if (element != null && !element.getText().isEmpty()) {
			                 String text = element.getText();
			                 return isValidValue(text) ? text.replace("₹", "").replace("MRP:", "").trim() : "NA";
			             }

			             // If neither section has a valid MRP, return "NA"
			             return "NA";

			         } catch (Exception e) {
			             // Handle any other exceptions by returning "NA"
			             return "NA";
			         }
			     }



			     private static String extractSP(WebDriver driver) {
			         try {
			             List<WebElement> elements = driver.findElements(By.xpath("//span[@class='price-item price-item--sale price-item--last']"));

			             if (elements.isEmpty()) {
			                 elements = driver.findElements(By.xpath("//span[@class='price-item price-item--discounted']"));
			             }

			             if (!elements.isEmpty()) {
			                 String text = elements.get(0).getText();
			                 return isValidValue(text) ? text.replace("₹", "").replace("MRP:", "").trim() : "NA";
			             } else {
			                 return "NA";
			             }
			         } catch (Exception e) {
			             return "NA";
			         }
			     }

			     private static String extractOffer(WebDriver driver) {
			         try {
			             WebElement offerElement = driver.findElement(By.xpath("//span[@class='percent_discount']"));
			             return offerElement.getText();

			         } catch (Exception e) {
			             return "NA";  // If offer not found
			         }
			     }

			     private static String extractProductId(String url) {
			         // Extract the product ID from the URL if applicable
			         // Placeholder for extraction logic based on URL pattern
			         return url;
			     }

			     private static boolean isValidValue(String value) {
			         return value != null && !value.isEmpty() && !value.equals("NA");
			     }

			     private static String extractWeight(List<String> UOM, int i) {
			         // Return the UOM directly from the input data list based on the current index `i`
			         return UOM.get(i);  // Return the UOM for the current product
			     }

			     private static void takeScreenshot(WebDriver driver, String productId) {
			         try {
			             TakesScreenshot ts = (TakesScreenshot) driver;
			             File sourceFile = ts.getScreenshotAs(OutputType.FILE);
			             String timestamp = new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss").format(new Date());
			             String screenshotPath = ".\\Screenshot\\" + "HeadsUpForTails_" + productId + "_" + timestamp + ".png";
			             File destinationFile = new File(screenshotPath);
			             FileUtils.copyFile(sourceFile, destinationFile);
			             System.out.println("Screenshot saved for Product ID: " + productId);
			         } catch (Exception e) {
			             System.out.println("Failed to capture screenshot: " + e.getMessage());
			         }
			     }
			    }




