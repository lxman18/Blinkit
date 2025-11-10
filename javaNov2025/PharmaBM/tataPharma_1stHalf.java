package PharmaSunday;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
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
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.testng.Assert;

import java.util.Calendar;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

public class tataPharma_1stHalf {
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
		nextRunTime.set(Calendar.HOUR_OF_DAY, 15);
		nextRunTime.set(Calendar.MINUTE, 0);
		nextRunTime.set(Calendar.SECOND, 0);

		long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		if (initialDelay < 0) {
			initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		}

		scheduler.scheduleAtFixedRate(() -> {
			try {
				System.out.println("Starting web scraping task...");
				tataPharma_1stHalf.runWebScraping();
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
			Sheet urlsSheet = urlsWorkbook.getSheet("TATA 1mg-1");
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
					
					
					  String pageTitle = driver.getTitle();
					  
					// Add this check here
					  if (pageTitle.startsWith("Online Pharmacy India")) {
					    Row resultRow = resultsSheet.createRow(rowIndex++);
					    
					    resultRow.createCell(0).setCellValue(id);
						resultRow.createCell(1).setCellValue(city);
						resultRow.createCell(2).setCellValue(name); 
						resultRow.createCell(3).setCellValue(size);
						resultRow.createCell(4).setCellValue("NA");
						resultRow.createCell(5).setCellValue("NA");
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
						resultRow.createCell(17).setCellValue(" ");
						resultRow.createCell(18).setCellValue(namecheck);
					    System.out.println("Found leading platform text, skipping extraction for URL: " + url);
					    continue;
					}

					try {

						WebElement nameElement = driver.findElement(By.xpath("//h1"));
						newName = nameElement.getText();
						System.out.println(newName);
					}

					catch(NoSuchElementException e) {

						WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/div/div/div/div/div[2]/div[3]/div[1]/div[1]/h1"));
						newName = nameElement.getText();
						System.out.println(newName);

					}
					System.out.println("headercount = " + headercount);

					headercount++;
					
					Thread.sleep(5000);
					
					
					//sp
					try {
						WebElement sp = driver.findElement(By.xpath("//div[@id='normal-pricebox']//span[1]"));
						originalSp1 = sp.getText();
						spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "").replace("Discounted Price: ", "");
					
					}
					catch (Exception as) {
						try {
							WebElement sp = driver.findElement(By.xpath("//div[contains(@class, 'PriceDetails__discount-div')]"));
							String fullText = sp.getText(); 
							
							Pattern pattern = Pattern.compile("\\d+(\\.\\d+)?"); 
							Matcher matcher = pattern.matcher(fullText);

							String priceOnly = "";
							if (matcher.find()) {
							    priceOnly = matcher.group(); // Extract the number like "320.5" or "119"
							}

							System.out.println("Extracted Price: " + priceOnly);
							spValue=priceOnly;
						}
						catch (Exception sdf) {
							try {
								WebElement sp = driver.findElement(By.xpath("//div[@class='DrugPriceBox__best-price___32JXw']"));
								originalSp1 = sp.getText();
								spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "").replace("Discounted Price: ", "");
							
							}catch (Exception sae) {
								try{
									WebElement sp = driver.findElement(By.xpath("//div[@class='DrugPriceBox__price___dj2lv']"));
									originalSp1 = sp.getText();
									spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "").replace("Discounted Price: ", "");
							}catch (Exception ew) {
								spValue="NA";
							}
						}
					}}
					 System.out.println(spValue);
					
					//mrp
					try {
						WebElement mrp = driver.findElement(By.xpath("//div[@id='normal-pricebox']//span[2]"));
						originalMrp1 = mrp.getText();
						mrpValue = originalMrp1.replace("₹", "");
						

					}catch (Exception sa) {
						try {
						WebElement mrp = driver.findElement(By.xpath("//span[@class='DiscountDetails__discount-price___Mdcwo']"));
						originalMrp1 = mrp.getText();
						mrpValue = originalMrp1.replace("₹", "");
						}catch (Exception e) {
							try {
								WebElement mrp = driver.findElement(By.xpath("//span[@class='DrugPriceBox__slashed-price___2UGqd']"));
								originalMrp1 = mrp.getText();
								mrpValue = originalMrp1.replace("₹", "");
							}catch (Exception em) {
								mrpValue=spValue;
							}
						}
					}
					System.out.println(mrpValue);
					

					//Out Of Stocks
					if(url.contains("NA")){
						String result = "NA";
					}	

					int result=1;
					try {
					WebElement cart=driver.findElement(By.xpath("//div[contains(text(),'Add to cart')]"));
					if(cart.isDisplayed() && cart.isEnabled()) {
						result=1;
					}}catch (Exception e){
						result=0;
					}
					NewAvailability1 = String.valueOf(result);
					System.out.println(NewAvailability1);

					// OFFER

					if(mrpValue.equals(spValue)){
						offerValue = "NA";
					}
					else {
						try {
							WebElement offer = driver.findElement(By.className("DiscountDetails__discount-percent___IfDdk"));
							String offer1 = offer.getText();
							String visibleText = (String) ((JavascriptExecutor) driver).executeScript(
								    "return arguments[0].lastChild.textContent.trim();", offer);

							String offer2 = visibleText.replace("% off", "% Off").replace("Discount Percentage: ", "").trim();
							offerValue = offer2;
						}
						catch(Exception e) {
							try {
							WebElement offer = driver.findElement(By.xpath("//div[@id='normal-pricebox']//span[3]"));
							String offer1 = offer.getText();
							String offer2 = offer1.replace("% off", "% Off").replace("Discount Percentage: ", "").trim();

							offerValue = offer2;
						}catch (Exception ess) {
							WebElement offer = driver.findElement(By.xpath("//span[@class='DrugPriceBox__slashed-percent___G92cz']"));
							String offer1 = offer.getText();
							String offer2 = offer1.replace("% off", "% Off").replace("Discount Percentage: ", "").trim();

							offerValue = offer2;
						}
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
				String outputFilePath = ".\\Output\\TAta1mg_FirstHalf_OutputData_" + timestamp + ".xlsx";


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
