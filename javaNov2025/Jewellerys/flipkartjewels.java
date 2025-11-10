package jewellery;

import java.io.FileOutputStream;
import java.lang.Thread;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class flipkartjewels {

    public static void main(String[] args) throws InterruptedException {
        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless=new");
        options.addArguments("--disable-gpu");
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.managed_default_content_settings.images", 2);
        prefs.put("profile.managed_default_content_settings.stylesheets", 2);
        prefs.put("profile.managed_default_content_settings.fonts", 2);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        
        String baseUrl = "https://www.flipkart.com/jewellery/gemstones-coins-bars/pr?sid=mcr%2C73x&p%5B%5D=facets.weight%255B%255D%3D10%2Bg&otracker=categorytree&otracker=nmenu_sub_Women_0_Coins+and+Bars&p%5B%5D=facets.material%255B%255D%3DSilver";
        
        // Set up Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Page Number");
        
        headerRow.createCell(1).setCellValue("Brand");
        headerRow.createCell(2).setCellValue("Category");
        headerRow.createCell(3).setCellValue("URL");
        headerRow.createCell(4).setCellValue("Name");
        headerRow.createCell(5).setCellValue("MRP");
        headerRow.createCell(6).setCellValue("SP");
        headerRow.createCell(7).setCellValue("UOM");
        headerRow.createCell(8).setCellValue("Availability");
        headerRow.createCell(9).setCellValue("Offer");
       
        // Shutdown hook to save progress on interruption
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        final String interruptFilePath = ".\\Output\\Flipkart_interrupt_" + timestamp + ".xlsx";
        final XSSFWorkbook finalWorkbook = workbook; // For lambda access
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            try (FileOutputStream out = new FileOutputStream(interruptFilePath)) {
                finalWorkbook.write(out);
                System.out.println("Interrupted save: " + interruptFilePath);
            } catch (Exception e) {
                System.out.println("Error saving interrupt: " + e.getMessage());
            } finally {
                try {
                    finalWorkbook.close();
                } catch (Exception ignored) {}
            }
        }));
        
        int page = 1;
        boolean hasNextPage = true;
        int rowNum = 1; // Start after header
        String pincode = "110015"; // Hardcoded pincode, change as needed
        
        while (hasNextPage) {
            String currentUrl = (page == 1) ? baseUrl : baseUrl + "&page=" + page;
            
            Thread.sleep(3000);
            driver.get(currentUrl);
            driver.manage().window().maximize();
            
            // Wait for products to load
            wait.until(ExpectedConditions.numberOfElementsToBeMoreThan(By.xpath("//div[@class='wvIX4U']"), 0));
            
            List<WebElement> products = driver.findElements(By.xpath("//div[@class='wvIX4U']"));
            
            if (products.isEmpty()) {
                hasNextPage = false;
                break;
            }
            
            String currentHandle = driver.getWindowHandle();
            
            for (WebElement product : products) {
                try {
                    product.click();
                    
                    // Wait for new tab to open
                    Set<String> handles = new HashSet<>(driver.getWindowHandles());
                    String newHandle = null;
                    while (newHandle == null) {
                        Set<String> updatedHandles = driver.getWindowHandles();
                        for (String h : updatedHandles) {
                            if (!h.equals(currentHandle)) {
                                newHandle = h;
                                break;
                            }
                        }
                        Thread.sleep(5000); // Small wait to allow tab to open
                    }
                    Thread.sleep(3000);
                    driver.switchTo().window(newHandle);
                    
//                    // Set pincode for availability
//                    try {
//                        try {
//							WebElement pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//							    By.xpath("//input[contains(@placeholder, 'Enter pincode') or contains(@placeholder, 'Delivery Location')]")));
//							pincodeInput.click();
//
//							pincodeInput.clear();
//							pincodeInput.sendKeys(pincode);
//						} catch (Exception e) {
//							WebElement pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//								    By.xpath("//div[@class='JqZtEs']")));
//								pincodeInput.clear();
//								pincodeInput.click();
//								pincodeInput.sendKeys(pincode);
//						}
//                        
//                        // Click check button
//                        WebElement checkButton = driver.findElement(By.xpath("//span[contains(text(), 'Check')] | //button[contains(text(), 'Check')]"));
//                        checkButton.click();
//                        
//                        // Wait for availability update
//                        Thread.sleep(3000);
//                    } catch (Exception e) {
//                        System.out.println("Pincode setting skipped: " + e.getMessage());
//                    }
//                    
                    // Get URL
                    String productUrl = driver.getCurrentUrl();
                    
                    // Wait for product details to load
                    wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("h1")));
                    
                    int headercount = 1;
                    
                    
                    String Brand1 = "";
                    try {
                        WebElement Brand = driver.findElement(By.xpath("//span[@class='mEh187']"));
                        Brand1 = Brand.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Brand = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[3]/div[1]/div[2]/div[2]/div/div[1]/h1/span[1]"));
                            Brand1 = Brand.getText();
                        } catch (NoSuchElementException ex) {
                        	Brand1 = "NA";
                        }
                    }
                    // Scrape details
                    String newName = "";
                    try {
                        WebElement nameElement = driver.findElement(By.xpath("//span[@class='VU-ZEz']"));
                        newName = nameElement.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement nameElement = driver.findElement(By.xpath("//h1[@class='_6EBuvT']//span[@class='VU-ZEz']"));
                            newName = nameElement.getText();
                        } catch (NoSuchElementException ex) {
                            newName = "NA";
                        }
                    }
                    System.out.println("Product Name: " + newName);
                    
                    String spValue = "NA";
                    String originalSp1 = "";
                    String mrpValue = "NA";
                    String originalMrp1 = "";
                    try {
                        WebElement sp = driver.findElement(By.xpath("//div[@class='yRaY8j A6+E6v']"));
                        originalSp1 = sp.getText();
                        spValue = originalSp1.replace("₹", "").replace(",", "");
                    } catch (NoSuchElementException s) {
                        try {
                            WebElement sp = driver.findElement(By.xpath("//div[@class='hl05eU']//div[@class='yRaY8j A6+E6v']"));
                            originalSp1 = sp.getText();
                            spValue = originalSp1.replace("₹", "").replace(",", "");
                        } catch (Exception t) {
                        	 spValue=mrpValue;
                        }
                    }
                    System.out.println("Selling Price: " + spValue);
                    
                   
                    try {
                        WebElement mrp = driver.findElement(By.xpath("//div[@class='Nx9bqj CxhGGd']"));
                        originalMrp1 = mrp.getText();
                        mrpValue = originalMrp1.replace("₹", "").replace(",", "");
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement mrp = driver.findElement(By.xpath("//div[@class='hl05eU']//div[@class='Nx9bqj CxhGGd']"));
                            originalMrp1 = mrp.getText();
                            mrpValue = originalMrp1.replace("₹", "").replace(",", "");
                        } catch (Exception t) {
                            mrpValue = spValue;
                        }
                    }
                    System.out.println("MRP: " + mrpValue);
                    
                    String offerValue = "NA";
                    if (!mrpValue.equals(spValue) && !"NA".equals(mrpValue) && !"NA".equals(spValue)) {
                        try {
                            WebElement offer = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[3]/div[1]/div[2]/div[2]/div/div[3]/div[1]/div/div[3]"));
                            String originalOffer = offer.getText();
                            offerValue = originalOffer.replace("off", "Off");
                        } catch (Exception e) {
                            try {
                                WebElement offer = driver.findElement(By.xpath("//div[@class='UkUFwK WW8yVX dB67CR']//span"));
                                String originalOffer = offer.getText();
                                offerValue = originalOffer.replace("off", "Off");
                            } catch (Exception h) {
                                offerValue = "NA";
                            }
                        }
                    }
                    System.out.println("Offer: " + offerValue);
                    
                    //UOM (Weight)
                    String uomValue = "NA";
                    try {
                        WebElement uom = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//li[@class='aJWdJI dpZEpc']//a)[1]")));
                        uomValue = uom.getText().trim();
                    } catch (Exception e) {
                        // Alternative XPath if needed
                        try {
                            WebElement uom = driver.findElement(By.xpath("//*[@id=\"swatch-0-weight\"]/a"));
                            uomValue = uom.getText().trim();
                        } catch (Exception ex) {
                            uomValue = "NA";
                        }
                    }
                    System.out.println("UOM (Weight): " + uomValue);
                    
                    // Availability
                    int result = 1;
                    String[] textsToCheck = { "Currently Unavailable", "Currently out of stock for", "Sold Out", "NOTIFY ME" };
                    String pageSource = driver.getPageSource();
                    boolean outOfStock = false;
                    for (String text : textsToCheck) {
                        if (pageSource.toLowerCase().contains(text.toLowerCase())) {
                            outOfStock = true;
                            break;
                        }
                    }
                    if (outOfStock) {
                        result = 0;
                    } else {
                        try {
                            WebElement atc = null;
                            try {
                                atc = driver.findElement(By.xpath("//button[@class='QqFHMw vslbG+ In9uk2']"));
                            } catch (Exception eeea) {
                                try {
                                    atc = driver.findElement(By.xpath("//button[@class='QqFHMw vslbG+ _3Yl67G _7Pd1Fp']"));
                                } catch (Exception ignored) {
                                    // atc remains null
                                }
                            }
                            if (atc != null && atc.isDisplayed() && atc.isEnabled()) {
                                result = 1;
                            } else {
                                result = 0;
                            }
                        } catch (Exception eee) {
                            result = 0;
                        }
                    }
                    
                    String availability = String.valueOf(result);
                    System.out.println("Availability: " + availability);
                    
                    try {
						System.out.println("Scraped Url"+productUrl);
					} catch (Exception esd) {
						System.out.println("Failed Url"+productUrl);
					}
                    
                    System.out.println("===================================== product count :"+ headercount++ +"=====================================================");
                    // Write to Excel
                    Row dataRow = sheet.createRow(rowNum++);
                    dataRow.createCell(0).setCellValue(page);
                    dataRow.createCell(1).setCellValue(Brand1);
                    dataRow.createCell(2).setCellValue("10 Gm Silver");
                    
                    dataRow.createCell(3).setCellValue(productUrl);
                    dataRow.createCell(4).setCellValue(newName);
                    dataRow.createCell(5).setCellValue(mrpValue);
                    dataRow.createCell(6).setCellValue(spValue);
                    dataRow.createCell(7).setCellValue(uomValue);
                    dataRow.createCell(8).setCellValue(availability);
                    dataRow.createCell(9).setCellValue(offerValue);
                    
                    
                    System.out.println("---");
                    
                } catch (Exception eaq) {
                    System.out.println("Error processing product: " + eaq.getMessage());
                } finally {
                    // Close the new tab and switch back
                    try {
                        driver.close();
                        driver.switchTo().window(currentHandle);
                    } catch (Exception ex) {
                        // Ignore if already closed
                    }
                }
            }
            
            // Save after each page to preserve progress
            try {
                String progressFilePath = ".\\Output\\Flipkart_progress_page_" + page + "_" + timestamp + ".xlsx";
                try (FileOutputStream out = new FileOutputStream(progressFilePath)) {
                    workbook.write(out);
                }
                System.out.println("Progress saved after page " + page + ": " + progressFilePath);
            } catch (Exception e) {
                System.out.println("Error saving progress: " + e.getMessage());
            }
            
            page++;
            // Optional: Add max page limit, e.g., if (page > 50) { hasNextPage = false; }
        }
        
        // Final save
        String outputFilePath = ".\\Output\\Flipkart_output_" + timestamp + ".xlsx";
        try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
            workbook.write(outFile);
            System.out.println("Final output file saved: " + outputFilePath);
        } catch (Exception e) {
            System.out.println("Error saving final Excel: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (Exception e) {
                // Ignore
            }
        }
        
        driver.quit();
    }
}