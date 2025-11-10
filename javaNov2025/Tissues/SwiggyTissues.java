
package AssortmentTissues;
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
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SwiggyTissues {

    // Class to hold product details
    static class Product {
        String name;
        String mrp;
        String sp;
        String uom;
        int availability;
        String offer;
        String color;
        String boxContents;
        String material;
        String packSize;
        String usage;
        String url;

        Product(String name, String mrp, String sp, String uom, int availability, String offer, String color,
                String boxContents, String material, String packSize, String usage, String url) {
            this.name = name;
            this.mrp = mrp;
            this.sp = sp;
            this.uom = uom;
            this.availability = availability;
            this.offer = offer;
            this.color = color;
            this.boxContents = boxContents;
            this.material = material;
            this.packSize = packSize;
            this.usage = usage;
            this.usage = url;
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

            // Perform the search for "Kitchen Disposables"
            System.out.println("Performing search for 'Kitchen Disposables'...");
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@data-testid='search-container']"))).click();

            WebElement searchInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@type='search']")));
            searchInput.sendKeys("Kitchen Disposables");
            searchInput.sendKeys(Keys.ENTER);
            Thread.sleep(3000); // Wait for search results to load

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
                        System.out.println("Error page detected after search. Attempting retry " + (retryCount + 1));
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
                        System.out.println("Search results page loaded successfully.");
                        break;
                    }
                } catch (Exception e) {
                    System.out.println("Exception during search results load: " + e.getMessage());
                    retryCount++;
                    driver.navigate().refresh();
                    Thread.sleep(2000);
                }
            }

            if (retryCount == maxRetries) {
                System.out.println("Max retries reached after search. Could not load the page successfully.");
                driver.quit();
                return;
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
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class, 'gmdlfz')] | //div[contains(@class, 'sc-gEvEer')]")));

                System.out.println("--- Product " + (j + 1) + " ---");

                // Extract product details
                
                String name= "";
                String mrp = "";
                String sp = "";
                String uom  = "";
                int availability;
                String offer = "";
                String color = "";
                String boxContents = "";
                String material = "";
                String packSize = "";
                String usage = "";
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
					System.out.println("SP: " + uom);
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

                
				try {
					color = extractText(driver, wait, By.xpath("(//div[contains(.,'Colour')]//div[@class='_1aE0r'])[1]"),
					    By.xpath("(//div[contains(.,'Colour')]//div[@class='_1aE0r']//div)[1]"), "NA");
					System.out.println("Color: " + color);
				} catch (Exception e) {
					color ="NA";
					System.out.println("Color: " + color);
				}

                
				try {
					boxContents = extractText(driver, wait, By.xpath("(//div[contains(.,'Box Contents')]//div[@class='_1aE0r'])[2]"),
					    By.xpath("(//div[contains(.,'Box Contents')]//div[@class='_1aE0r']//div)[2]"), "NA");
					System.out.println("Box Contents: " + boxContents);
				} catch (Exception e) {
					boxContents = "NA";
					System.out.println("Color: " + color);
				}

                
				try {
					material = extractText(driver, wait, By.xpath("(//div[contains(.,'Material')]//div[contains(@class,'_1aE0r')])[3]"),
					    By.xpath("(//div[contains(.,'Material')]//div[contains(@class,'_1aE0r')]//div)[3]"), "NA");
					System.out.println("Material: " + material);
				} catch (Exception e) {
					material ="NA";
					System.out.println("Material: " + material);
				}

                
				try {
					packSize = extractText(driver, wait, By.xpath("(//div[contains(.,'Pack Size')]//div[contains(@class,'_1aE0r')])[4]"),
					    By.xpath("(//div[contains(.,'Pack Size')]//div[contains(@class,'_1aE0r')]//div)[4]"), "NA");
					System.out.println("Pack Size: " + packSize);
				} catch (Exception e) {
					packSize ="NA";
					System.out.println("Pack Size: " + packSize);
				}

                
				try {
					usage = extractText(driver, wait, By.xpath("(//div[contains(.,'Pack Size')]//div[contains(@class,'_1aE0r')])[5]"),
					    By.xpath("(//div[contains(.,'Pack Size')]//div[contains(@class,'_1aE0r')]//div)[5]"), "NA");
					System.out.println("Usage: " + usage);
				} catch (Exception e) {
					usage ="NA";
					System.out.println("Pack Size: " + usage);
					
				}
				
				

                // Add product to list
                productList.add(new Product(name, mrp, sp, uom, availability, offer, color, boxContents, material, packSize, usage,url));

                // Navigate back
                driver.navigate().back();
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='_3Rr1X']")));
            }

            // Write to Excel
            writeToExcel(productList);

        } catch (Exception e) {
            System.out.println("An error occurred: " + e.getMessage());
        } finally {
            driver.quit();
        }
    }

    // Helper method to extract text with fallback
    private static String extractText(WebDriver driver, WebDriverWait wait, By primaryLocator, By fallbackLocator, String fieldName) {
        try {
            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(primaryLocator));
            return element.getText().isEmpty() ? "NA" : element.getText();
        } catch (Exception e) {
            try {
                WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(fallbackLocator));
                return element.getText().isEmpty() ? "NA" : element.getText();
            } catch (Exception ex) {
                System.out.println(fieldName + " not found.");
                return "NA";
            }
        }
    }

    // Method to write product data to Excel
    private static void writeToExcel(List<Product> productList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Kitchen Disposables");

        // Create header row
        String[] headers = {"Name", "MRP", "SP", "UOM", "Availability", "Offer", "Color", "Box Contents", "Material", "Pack Size", "Usage","URL"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Populate data rows
        for (int i = 0; i < productList.size(); i++) {
            Product product = productList.get(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(product.name);
            row.createCell(1).setCellValue(product.mrp);
            row.createCell(2).setCellValue(product.sp);
            row.createCell(3).setCellValue(product.uom);
            row.createCell(4).setCellValue(product.availability);
            row.createCell(5).setCellValue(product.offer);
            row.createCell(6).setCellValue(product.color);
            row.createCell(7).setCellValue(product.boxContents);
            row.createCell(8).setCellValue(product.material);
            row.createCell(9).setCellValue(product.packSize);
            row.createCell(10).setCellValue(product.usage);
            row.createCell(11).setCellValue(product.url);
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write to file
        
        try (FileOutputStream fileOut = new FileOutputStream(".\\Output\\Tissuse.xlsx")) {
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