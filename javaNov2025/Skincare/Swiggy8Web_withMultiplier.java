package blinkitTrial;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Swiggy8Web_withMultiplier {

	public static void main(String[] args) throws Exception{

		/*
		 * ChromeOptions options = new ChromeOptions();
		 * 
		 * // Stealth: Make it less detectable
		 * options.addArguments("--disable-blink-features=AutomationControlled");
		 * options.setExperimentalOption("excludeSwitches", new
		 * String[]{"enable-automation"});
		 * options.setExperimentalOption("useAutomationExtension", false);
		 * 
		 * // Use mobile user-agent options.
		 * addArguments("user-agent=Mozilla/5.0 (Linux; Android 10; SM-G975F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Mobile Safari/537.36"
		 * );
		 * 
		 * // Optional: Set geolocation permission to allow Map<String, Object> prefs =
		 * new HashMap<>();
		 * prefs.put("profile.default_content_setting_values.geolocation", 1); // 1 =
		 * allow options.setExperimentalOption("prefs", prefs);
		 */

		WebDriver driver = new ChromeDriver();
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

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
		String webUom = " ";
		double multiplier = 0.0;

		try {
			// Read URLs from Excel file
			String filePath = ".\\InputData\\Website 8 Input data.xlsx"; //  Website 8 Input data.xlsx
			FileInputStream file = new FileInputStream(filePath);
			Workbook urlsWorkbook = new XSSFWorkbook(file);
			Sheet urlsSheet = urlsWorkbook.getSheet("Swiggy"); // Swiggy
			int rowCount = urlsSheet.getPhysicalNumberOfRows();

			List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
					uRL = new ArrayList<>();

			// Extract URLs from Excel
			for (int i = 0; i < rowCount; i++) {
				Row row = urlsSheet.getRow(i);                               
				if (i == 0) {
					continue;
				}     

				String id = getCellValue(row.getCell(0));
				String city = getCellValue(row.getCell(1));
				String name = getCellValue(row.getCell(2));
				String size = getCellValue(row.getCell(3));
				String productCode = getCellValue(row.getCell(4));
				String url = getCellValue(row.getCell(5));

				inputPid.add(id);
				InputCity.add(city);
				InputName.add(name);
				InputSize.add(size);
				NewProductCode.add(productCode);
				uRL.add(url);
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

					if(i==0) {
						driver.get("https://www.swiggy.com/"); 
						Thread.sleep(3000);

						try {
							WebElement pin = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter your delivery location']"))); 
							pin.click();
							pin.sendKeys("110015");

//							WebElement type = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter your delivery location']")));
//							type.sendKeys("110015"); // 110015

							Thread.sleep(1000);

							WebElement drop = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='kuQWc'])[1]"))); 
							drop.click();     
						} 
						catch (Exception e) {
							WebElement pin = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[text()='Other']"))); 
							pin.click();

							WebElement type = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Search for area, street name..']")));
							type.sendKeys("110015");

							Thread.sleep(1000);

							WebElement drop = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='icon-location _2H6Kl'])[1]"))); 
							drop.click();    
							
							Thread.sleep(1000);
						}										 
					}
					driver.manage().window().maximize();
					driver.get(url);
					
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

					    System.out.println("Something went wrong found, URL skipped : " + url);
					    continue;  // skip to next URL in loop
					}

					
					} catch (NoSuchElementException e) {
						try {

							WebElement nameElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='sc-aXZVg aUBRn _1iFYi']")));
							newName = nameElement.getText();
							System.out.println(newName);
						}

						catch(NoSuchElementException ed) {
						}
						System.out.println("headercount = " + headercount);

						headercount++;
						
						//sp
						try {
							WebElement sp = driver.findElement(By.xpath("//div[@data-testid='item-offer-price']"));
							originalSp1 = sp.getText();
							spValue =  originalSp1.replace("₹", "");
							
							if(spValue.isEmpty() || spValue.equalsIgnoreCase("NaN")) {
								spValue = "NA";
							}
							System.out.println(spValue);
						}
						catch(Exception ae) {
					}
						// Mrp 
						try {
							WebElement mrp = driver.findElement(By.xpath("//div[@data-testid='item-mrp-price']"));
							originalMrp1 = mrp.getText();
							mrpValue = originalMrp1.replace("₹", "");
							
							if (mrpValue.isEmpty() || mrpValue.equalsIgnoreCase("NaN")) {
								mrpValue = "NA";
							}
							System.out.println(mrpValue);
							
						}
						catch(NoSuchElementException eh){ 
							try {
								WebElement mrp = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]"));
								originalMrp1 = mrp.getText();
								mrpValue = originalMrp1.replace("₹", "");
								
								if (mrpValue.isEmpty() || mrpValue.equalsIgnoreCase("NaN")) {
									mrpValue = "NA";
								}
								System.out.println(mrpValue);
							}
							catch (Exception S) {
								mrpValue=spValue;
							}
						}

						//uom  
						try {
							WebElement webUom1 = driver.findElement(By.xpath("//div[@class='sc-aXZVg kYaBqd _11EdJ']"));
							String webUom2 = webUom1.getText();
							webUom = webUom2;
							System.out.println(webUom);
						}
						catch (Exception u) {
							u.printStackTrace();
						}

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
							} catch (Exception ehf) {
								System.out.println("Error checking availability: " + e.getMessage());
								result = -1;
							}
						}

						// Assign final availability status
						NewAvailability1 = String.valueOf(result);

						// Multiplier 
						multiplier = calculateMultiplier(size, webUom);
						System.out.println(multiplier);
						

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
						resultRow.createCell(9).setCellValue(webUom);
						resultRow.createCell(10).setCellValue(multiplier);
						resultRow.createCell(11).setCellValue(NewAvailability1);
						resultRow.createCell(12).setCellValue(" ");
						resultRow.createCell(13).setCellValue(" ");
						resultRow.createCell(14).setCellValue(" ");
						resultRow.createCell(15).setCellValue(" ");
						resultRow.createCell(16).setCellValue(" ");
						resultRow.createCell(17).setCellValue(offerValue);                    

						System.out.println("Data extracted for URL: " + url);
					} }catch (Exception e) {
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
						resultRow.createCell(9).setCellValue(webUom);
						resultRow.createCell(10).setCellValue(multiplier);
						resultRow.createCell(11).setCellValue(NewAvailability1);
						resultRow.createCell(12).setCellValue(" ");
						resultRow.createCell(13).setCellValue(" ");
						resultRow.createCell(14).setCellValue(" ");
						resultRow.createCell(15).setCellValue(" ");
						resultRow.createCell(16).setCellValue(" ");
						resultRow.createCell(17).setCellValue(offerValue);                  

						System.out.println("Failed to extract data for URL: " + url);
					}
			}

			// Write results to Excel file
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            
            String outputFilePath = ".\\Output\\Swiggy_SkinCare" + timestamp + ".xlsx"; 
			FileOutputStream outFile = new FileOutputStream(outputFilePath);
			resultsWorkbook.write(outFile);
			outFile.close();

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (driver != null) {
				driver.quit();
				System.out.println("Scraping Completed.");
			}
		}
	}
	public static String getCellValue(Cell cell) {
		if (cell == null) return "";
		if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
		else if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
		return "";
	}

	private static double calculateMultiplier(String inputSize, String webUom) {
		try {
			// Clean and extract values
			inputSize = inputSize.toLowerCase().replaceAll("[^0-9.]", "").trim();
			webUom = webUom.toLowerCase().trim();

			double inputGrams = Double.parseDouble(inputSize);
			double outputGrams;

			if (webUom.contains("x")) {
				String[] parts = webUom.split("x");
				double base = Double.parseDouble(parts[0].replaceAll("[^0-9.]", "").trim());
				double times = Double.parseDouble(parts[1].trim());
				outputGrams = base * times;
			} else {
				outputGrams = Double.parseDouble(webUom.replaceAll("[^0-9.]", "").trim());
			}

			if (outputGrams == 0) return 0;

			double multiplier = inputGrams / outputGrams;
			return Math.round(multiplier * 100.0) / 100.0;

		} catch (Exception e) {
			e.printStackTrace();
			return 0;
		}
	}

}