package AssortmentTissues;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class FlipkartMinutes {

    public static void main(String[] args) {
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu");
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.managed_default_content_settings.images", 2);
        prefs.put("profile.managed_default_content_settings.stylesheets", 2);
        prefs.put("profile.managed_default_content_settings.fonts", 2);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(30));

        String baseUrl = "https://www.flipkart.com/search?q=tissues&otracker=search&otracker1=search&marketplace=HYPERLOCAL&as-show=on&as=off";
        String pincode = "700005";
        String location = "kolkata";

        // Create Output directory if not exists
        new File(".\\Output").mkdirs();

        // Set up Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Page Number");
        headerRow.createCell(1).setCellValue("BrandCategory");
        headerRow.createCell(2).setCellValue("Location");
        headerRow.createCell(3).setCellValue("URL");
        headerRow.createCell(4).setCellValue("Name");
        headerRow.createCell(5).setCellValue("MRP");
        headerRow.createCell(6).setCellValue("SP");
        headerRow.createCell(7).setCellValue("UOM");
        headerRow.createCell(8).setCellValue("Availability");
        headerRow.createCell(9).setCellValue("Offer");

        int page = 1;
        boolean hasNextPage = true;
        int rowNum = 1;

        while (hasNextPage) {
            String currentUrl = (page == 1) ? baseUrl : baseUrl + "&page=" + page;
            driver.get(currentUrl);
            driver.manage().window().maximize();
            // Set pincode for each page
            try {
                WebElement pincodeTrigger = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//div[contains(text(), 'Enter location manually')] | //div[contains(@class, 'JqZtEs')]")));
                pincodeTrigger.click();

                // Enter pincode
                WebElement pincodeInput;
                try {
                    pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("search")));
                } catch (Exception e) {
                    pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[contains(@placeholder, 'Enter')]")));
                }
                pincodeInput.clear();
                pincodeInput.sendKeys(pincode);

                // Select suggestion
                try {
                    WebElement suggestion = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("(//div[contains(@class, '_395Ij7')])[2]")));
                    suggestion.click();
                } catch (Exception e) {
                    System.out.println("No pincode suggestion found, proceeding...");
                }

                // Confirm pincode
                WebElement confirmButton = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//input[@class='sVrZD1'] | //button[contains(text(), 'Confirm')]")));
                confirmButton.click();

                wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.xpath("//span[contains(text(), '" + pincode + "')] | //div[contains(text(), 'Delivery')]")));
                System.out.println("Pincode " + pincode + " set successfully for page " + page);
            } catch (Exception e) {
                System.out.println("Error setting pincode: " + e.getMessage());
            }

            // Wait for products to load using a more generic product selector
            By productSelector = By.cssSelector("a[href*='/p/']"); // Adjusted to a more general selector

            try {
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(productSelector));
                System.out.println("Products loaded on page " + page);
            } catch (Exception e) {
                System.out.println("No products found on page " + page + ": " + e.getMessage());
                hasNextPage = false;
                break;
            }

            List<WebElement> products = driver.findElements(productSelector);

            if (products.isEmpty()) {
                System.out.println("No products found on page " + page);
                hasNextPage = false;
                break;
            }

            System.out.println("Found " + products.size() + " products on page " + page);

            for (int i = 0; i < products.size(); i++) {
                try {
                    // Re-fetch list to avoid stale references
                    products = driver.findElements(productSelector);
                    if (i >= products.size()) break;
                    WebElement product = products.get(i);
                    if (!product.isDisplayed()) continue;

                    // Store listing page URL
                    String listingPageUrl = driver.getCurrentUrl();

                    // Scroll to product and click
                    ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", product);
                    product.click();
                    wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("h1")));

                    String productUrl = driver.getCurrentUrl();
                    System.out.println("Navigated to product: " + productUrl);

                    
                    // Scrape Brand
                    String brand = "NA";
                    try {
                        WebElement brandElement = driver.findElement(By.xpath("(//div[@class='css-1rynq56'])[1]"));
                        brand = brandElement.getText().trim();
                    } catch (NoSuchElementException e) {
                        brand = "NA";
                    }
                    System.out.println("Brand: " + brand);

                    // Scrape Product Name
                    String name = "NA";
                    try {
                        WebElement nameElement = driver.findElement(By.xpath("(//span[@class='css-1qaijid'])[1]"));
                        name = nameElement.getText().trim();
                    } catch (NoSuchElementException e) {
                        name = "NA";
                    }
                    System.out.println("Product Name: " + name);

                    // Scrape Selling Price (SP)
                    String spValue = "NA";
                    try {
                        WebElement sp = driver.findElement(By.xpath("(//div[@class='css-1rynq56 r-11wrixw'])[2]"));
                        spValue = sp.getText().replace("₹", "").replace(",", "").trim();
                    } catch (NoSuchElementException e) {
                        spValue = "NA";
                    }
                    System.out.println("Selling Price: " + spValue);

                    // Scrape MRP
                    String mrpValue = "NA";
                    try {
                        WebElement mrp = driver.findElement(By.xpath("(//div[@class='css-1rynq56 r-11wrixw'])[1]"));
                        mrpValue = mrp.getText().replace("₹", "").replace(",", "").trim();
                    } catch (NoSuchElementException e) {
                        mrpValue = spValue;
                    }
                    System.out.println("MRP: " + mrpValue);

                    // Scrape Offer
                    String offerValue = "NA";
                    if (!mrpValue.equals(spValue) && !"NA".equals(mrpValue) && !"NA".equals(spValue)) {
                        try {
                            WebElement offer = driver.findElement(By.xpath("(//div[@class='css-1rynq56 r-11wrixw r-gy4na3'])[1]"));
                            offerValue = offer.getText().replace("off", "Off").trim();
                        } catch (Exception e) {
                            offerValue = "NA";
                        }
                    }
                    System.out.println("Offer: " + offerValue);

                    // Scrape UOM (Weight)
                    String uomValue = "NA";
                    try {
                        WebElement uom = wait.until(ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//span[@class='css-1qaijid r-1vgyyaa r-1b43r93 r-1rsjblm r-1x6tx0e r-1dupt2p']")));
                        uomValue = uom.getText().trim();
                    } catch (Exception e) {
                        uomValue = "NA";
                    }
                    System.out.println("UOM (Weight): " + uomValue);

                    // Check Availability
                    String availability = "1";
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
                        availability = "0";
                    }
                    System.out.println("Availability: " + availability);

                    // Write to Excel
                    Row dataRow = sheet.createRow(rowNum++);
                    dataRow.createCell(0).setCellValue(page);
                    dataRow.createCell(1).setCellValue(brand);
                    dataRow.createCell(2).setCellValue(location);
                    dataRow.createCell(3).setCellValue(productUrl);
                    dataRow.createCell(4).setCellValue(name);
                    dataRow.createCell(5).setCellValue(mrpValue);
                    dataRow.createCell(6).setCellValue(spValue);
                    dataRow.createCell(7).setCellValue(uomValue);
                    dataRow.createCell(8).setCellValue(availability);
                    dataRow.createCell(9).setCellValue(offerValue);

                    System.out.println("===================================== product count :" + (rowNum - 1) + "=====================================================");

                    // Navigate back to listing page
                    driver.get(listingPageUrl);
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(productSelector));

                } catch (StaleElementReferenceException e) {
                    System.out.println("Stale element for product " + (i + 1) + ", retrying...");
                    i--; // Retry current index
                } catch (Exception e) {
                    System.out.println("Error processing product " + (i + 1) + ": " + e.getMessage());
                }
            }

            // Save progress after each page
            try {
                String progressFilePath = ".\\Output\\FlipkartMinutes_Disposable&packaging_" + page + "_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";
                try (FileOutputStream out = new FileOutputStream(progressFilePath)) {
                    workbook.write(out);
                }
                System.out.println("Progress saved after page " + page + ": " + progressFilePath);
            } catch (Exception e) {
                System.out.println("Error saving progress: " + e.getMessage());
            }

            page++;
            if (page > 10) { // Adjust as needed
                hasNextPage = false;
            }
        }

        // Final save
        String outputFilePath = ".\\Output\\FlipkartMinutes_Tissues_kolkata" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";
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
            driver.quit();
        }
    }
}
