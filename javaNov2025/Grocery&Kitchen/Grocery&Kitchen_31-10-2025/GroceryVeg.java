package freshVeg;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class GroceryVeg {

    // Class to hold product details
    static class Product {
        String name;
        String mrp;
        String sp;
        String uom;
        int availability;
        String offer;
        String url;
        String l1 = "Grocery and Kitchen";
        String l2 = "Fresh Vegetables";
        String l3;

        Product(String name, String mrp, String sp, String uom, int availability, String offer, String url, String l3) {
            this.url = url;
            this.name = name;
            this.mrp = mrp;
            this.sp = sp;
            this.uom = uom;
            this.availability = availability;
            this.offer = offer;
            this.l3 = l3;
        }
    }

    public static void main(String[] args) {
        // List to store product details
        List<Product> productList = new ArrayList<>();

        // Set up Chrome options to avoid bot detection
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-blink-features=AutomationControlled");
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36");
        options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));
        options.setExperimentalOption("useAutomationExtension", false);

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        try {
            // Navigate to Swiggy Instamart main page to set location first
            driver.get("https://www.swiggy.com/instamart");
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");

            // Wait for page to load
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

            // Handle location selection (using a pincode for reliability)
            System.out.println("Handling location selection...");
            String city = "560001"; // Bangalore pincode; change to your preferred pincode
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@data-testid='search-location']"))).click();

            WebElement locationInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[contains(@placeholder,'Search for area')]")));
            locationInput.sendKeys(city);
            Thread.sleep(2000); // Wait for location suggestions to load

            // Select the first location suggestion
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[contains(@class,'icon-location-marker')]"))).click();
            Thread.sleep(2000); // Short wait for confirmation modal

            // Click Confirm button
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button/span[contains(text(),'Confirm')]"))).click();
            Thread.sleep(3000); // Wait for page to stabilize

            // Dismiss any re-check address tooltip if present
            try {
                String script = "let el = document.querySelector('div[data-testid=\"re-check-address-tooltip\"] > div[role=\"button\"]'); if (el) el.click();";
                js.executeScript(script);
                System.out.println("Dismissed re-check address tooltip if present.");
            } catch (Exception e) {
                System.out.println("No tooltip found or unable to dismiss: " + e.getMessage());
            }
            Thread.sleep(2000); // Additional wait for page load

            // Click on Fresh Vegetables main category
            System.out.println("Clicking on 'Fresh Vegetables' category...");
            wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//div[contains(text(),'Fresh Vegetables')]"))).click();
            Thread.sleep(5000);

            // Handle potential "Something went wrong" page
            int maxRetries = 5;
            int retryCount = 0;

            while (retryCount < maxRetries) {
                try {
                    wait.until(ExpectedConditions.or(
                        ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@class='_3Rr1X']")),
                        ExpectedConditions.presenceOfElementLocated(By.xpath("//*[contains(text(), 'Something went wrong')]"))
                    ));

                    if (driver.findElements(By.xpath("//*[contains(text(), 'Something went wrong')]")).size() > 0) {
                        System.out.println("Error page detected. Attempting retry " + (retryCount + 1));
                        try {
                            WebElement tryAgainButton = wait.until(ExpectedConditions.elementToBeClickable(
                                By.xpath("//button[contains(text(), 'Try Again') or contains(text(), 'Retry') or contains(text(), 'Refresh')]")
                            ));
                            tryAgainButton.click();
                            Thread.sleep(2000);
                            retryCount++;
                        } catch (Exception e) {
                            System.out.println("Try Again button not found. Refreshing page...");
                            driver.navigate().refresh();
                            retryCount++;
                        }
                    } else {
                        System.out.println("Category page loaded successfully.");
                        break;
                    }
                } catch (Exception e) {
                    System.out.println("Exception during page load: " + e.getMessage());
                    retryCount++;
                    driver.navigate().refresh();
                    Thread.sleep(2000);
                }
            }

            if (retryCount == maxRetries) {
                System.out.println("Max retries reached. Could not load the page successfully.");
                driver.quit();
                return;
            }

            // ---------- Sub-categories ----------
            List<WebElement> subcategories = wait.until(
                    ExpectedConditions.presenceOfAllElementsLocatedBy(
                            By.xpath("//div[@class='item-wrapper']")));

            System.out.println("Found " + subcategories.size() + " sub-categories.");

            for (int i = 0; i < subcategories.size(); i++) {
                try {
                    subcategories = driver.findElements(By.xpath("//div[@class='item-wrapper']"));
                    WebElement subCat = subcategories.get(i);
                    String subCatName = subCat.getText().trim();
                    if (subCatName.isEmpty()) continue;

                    System.out.println("\n=== " + subCatName + " ===");

                    scrollToElement(driver, subCat);
                    subCat.click();

                    JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
                    int last = 0;
                    while (true) {
                        jsExecutor.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                        List<WebElement> prods = driver.findElements(
                                By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
                        if (prods.size() <= last) break;
                        last = prods.size();
                        try {
                            jsExecutor.executeScript(
                                "let el=document.querySelector('div[data-testid=\"re-check-address-tooltip\"] button, div[role=\"button\"]');"
                              + "if(el)el.click();");
                        } catch (Exception ignored) {}
                    }

                    // Wait for products to be visible
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='_3Rr1X']")));

                    // Scroll to load all products
                    long lastHeight = (long) js.executeScript("return document.body.scrollHeight");
                    while (true) {
                        js.executeScript("window.scrollBy(0, 500)");
                        Thread.sleep(500);
                        long newHeight = (long) js.executeScript("return document.body.scrollHeight");
                        if (newHeight == lastHeight) {
                            break;
                        }
                        lastHeight = newHeight;
                    }

                    // Count products
                    List<WebElement> products = driver.findElements(By.xpath("//div[@class='_3Rr1X']"));
                    int count = products.size();
                    System.out.println("Number of products shown: " + count);

                    // Iterate through products
                    for (int j = 0; j < count; j++) {
                        products = driver.findElements(By.xpath("//div[@class='_3Rr1X']"));
                        WebElement product = products.get(j);
                        wait.until(ExpectedConditions.elementToBeClickable(product)).click();
                        Thread.sleep(6000);
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class, 'gmdlfz')] | //div[contains(@class, 'sc-gEvEer')]")));

                        System.out.println("--- Product " + (j + 1) + " ---");

                        // Extract product details
                        
                        String name= "";
                        String mrp = "";
                        String sp = "";
                        String uom  = "";
                        int availability;
                        String offer = "";
                        Thread.sleep(3000);
                        String url = driver.getCurrentUrl();
                       
                        
                         try {
    						name = extractText(driver, wait, By.xpath("//div[contains(@class, 'gmdlfz')] | //div[contains(@class, 'sc-gEvEer gmdlfz')]"),
    						    By.xpath("//div[contains(@class, 'sc-gEvEer gmd')]"), "NA");
    						System.out.println("Name: " + name);
    					} catch (Exception e) {
    						name = "NA";
    						System.out.println("Name: " + name);
    					}

                         try {
    						mrp = extractText(driver, wait, By.xpath("//div[contains(@class, 'gKjjWC')] | //div[contains(@class, 'sc-gEvEer gKjjWC')]"),
    						    By.xpath("(//div[contains(@class, 'sc-gEvEer gKjj')])[1]"), "NA");
    						System.out.println("MRP: " + mrp);
    					} catch (Exception e) {
    						mrp = sp;
    						System.out.println("MRP: " + mrp);
    					}

                         try {
    						sp = extractText(driver, wait, By.xpath("//div[contains(@class, 'iQcBUp')] | //div[contains(@class, 'sc-gEvEer iQcBUp')]"),
    						    By.xpath("(//div[contains(@class, 'sc-gEvEer iQcB')])[1]"), "NA");
    						System.out.println("SP: " + sp);
    					} catch (Exception e) {
    						sp = "NA";
    						System.out.println("SP: " + sp);
    					}

                         try {
    						uom = extractText(driver, wait, By.xpath("//div[contains(@class, 'ymEfJ')] | //div[contains(@class, 'sc-gEvEer ymEfJ')]"),
    						    By.xpath("(//div[contains(@class, 'sc-gEvEer ymEfJ')])[1]"), "NA");
    						System.out.println("UOM: " + uom);
    					} catch (Exception e) {
    						uom = "NA";
    						System.out.println("UOM: " + uom);
    					}

                        
    					try {
    						availability = driver.findElements(By.xpath("//button[contains(@class, '_1Imv1 _1L7yC _2rSIx')] | //button[contains(text(), 'Add') or contains(text(), 'ADD')]")).size() > 0 ? 1 : 0;
    						System.out.println("Availability: " + availability);
    					} catch (Exception e) {
    						availability =0;
    						System.out.println("Availability: " + availability);
    						
    					}

                        
    					try {
    						offer = extractText(driver, wait, By.xpath("//div[contains(@class, 'bsYAwc')] | //div[contains(@class, 'sc-gEvEer bsYAwc')]"),
    						    By.xpath("(//div[contains(@class, 'sc-gEvEer bsY')])[1]"), "NA");
    						System.out.println("Offer: " + offer);
    					} catch (Exception e) {
    						offer ="NA";
    						System.out.println("Offer: " + offer);
    					}
    					System.out.println("Product Url:"+ url);
                        
    					
    					

                        // Add product to list
                        productList.add(new Product(name, mrp, sp, uom, availability, offer, url, subCatName));

                        // Navigate back
                        driver.navigate().back();
                        wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='_3Rr1X']")));
                    }

                } catch (Exception e) {
                    System.out.println("Error processing subcategory: " + e.getMessage());
                }
            }

            // Write to Excel
            writeToExcel(productList);

        } catch (Exception e) {
            System.out.println("An error occurred: " + e.getMessage());
        } finally {
            driver.quit();
        }
    }

    // Helper method to scroll to element
    private static void scrollToElement(WebDriver driver, WebElement element) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].scrollIntoView(true);", element);
    }

    // Helper method to extract text with fallback
    private static String extractText(WebDriver driver, WebDriverWait wait, By primaryLocator, By fallbackLocator, String defaultValue) {
        try {
            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(primaryLocator));
            return element.getText().isEmpty() ? defaultValue : element.getText();
        } catch (Exception e) {
            try {
                WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(fallbackLocator));
                return element.getText().isEmpty() ? defaultValue : element.getText();
            } catch (Exception ex) {
                return defaultValue;
            }
        }
    }

    // Method to write product data to Excel
    private static void writeToExcel(List<Product> productList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Fresh Vegetables");

        // Create header row
        String[] headers = {"L1","L2","L3","URL","Name", "MRP", "SP", "UOM", "Availability", "Offer"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Populate data rows
        for (int i = 0; i < productList.size(); i++) {
            Product product = productList.get(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(product.l1);
            row.createCell(1).setCellValue(product.l2);
            row.createCell(2).setCellValue(product.l3);
            row.createCell(3).setCellValue(product.url);
            row.createCell(4).setCellValue(product.name);
            row.createCell(5).setCellValue(product.mrp);
            row.createCell(6).setCellValue(product.sp);
            row.createCell(7).setCellValue(product.uom);
            row.createCell(8).setCellValue(product.availability);
            row.createCell(9).setCellValue(product.offer);
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write to file
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        try (FileOutputStream fileOut = new FileOutputStream(".\\Output\\Fresh_Vegetables_"+timestamp+".xlsx")) {
            workbook.write(fileOut);
            System.out.println("===============Excel file created successfully: "+ fileOut +"===============");
        } catch (IOException e) {
            System.out.println("Error writing to Excel file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing workbook: " + e.getMessage());
            }
        }
    }
}