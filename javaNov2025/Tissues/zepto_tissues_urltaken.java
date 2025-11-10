package AssortmentTissues;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
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
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class dummy {

    public static void main(String[] args) throws InterruptedException {

        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless=new");
        options.addArguments("--disable-gpu");
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.managed_default_content_settings.images", 2);
        prefs.put("profile.managed_default_content_settings.stylesheets", 2);
        prefs.put("profile.managed_default_content_settings.fonts", 2);
        options.setExperimentalOption("prefs", prefs);
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/129.0.0.0");

        String headline = null;
        String spValue = "";
        String mrpValue = null;
        String offerValue = "NA";
        String newName = null;
        String uomValue = "NA";
        String brand = "NA";
        String NewAvailability1 = " ";
        int rowNum = 1;
        int headercount = 1;
        int page = 1;

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

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
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());

        final XSSFWorkbook finalWorkbook = workbook;
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            try {
                FileOutputStream out = new FileOutputStream("Zepto_" + timestamp + ".xlsx");
                finalWorkbook.write(out);
                out.close();
            } catch (Exception e) {
                System.out.println("Error saving on interrupt: " + e.getMessage());
            }
        }));

        String baseUrl = "https://www.zeptonow.com/cn/home-needs/tissues-disposables/cid/ac7b3ee1-98cb-48b8-8c62-27f47b4185a2/scid/9d0a7e11-73b6-4cbb-b00b-108b2f65e59f";
        driver.get(baseUrl);
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(200));
        
Thread.sleep(3000);
       

        // Wait for the page to load
        wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")));

        // Get headline (category)
        try {
            WebElement headli = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@class='c5SZXs ccdFPa']")));
            headline = headli.getText();
            System.out.println("Category: " + headline);
        } catch (Exception e) {
            headline = "Tissues & Disposables";
            System.out.println("Default Category: " + headline);
        }

        // Handle infinite scroll to collect all product URLs
        Set<String> productUrlSet = new HashSet<>();
        int previousSize = 0;
        int noChangeCount = 0;
        int maxNoChange = 5; // Increased to allow more cycles
        int cycle = 0;
        long previousHeight = (Long) ((JavascriptExecutor) driver).executeScript("return document.body.scrollHeight");

        System.out.println("Starting scroll cycle. Initial height: " + previousHeight);

        while (cycle < 50) {
            cycle++;

            // Scroll incrementally to bottom
            long currentHeight = (Long) ((JavascriptExecutor) driver).executeScript("return document.body.scrollHeight");
            long pos = 0;
            long increment = 300;
            while (pos < currentHeight) {
                ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, " + pos + ");");
                pos += increment;
                Thread.sleep(400);
                // Check for new links during scroll
                try {
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//a[contains(@href, '/pn/')]")));
                } catch (Exception e) {
                    System.out.println("Cycle " + cycle + ": No new links during scroll, continuing...");
                }
            }

            // Scroll back to top to trigger lazy loads
            ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, 0);");
            Thread.sleep(500);

            // Wait for content to load
            try {
                wait.until(ExpectedConditions.or(
                    ExpectedConditions.jsReturnsValue("return document.readyState === 'complete';"),
                    ExpectedConditions.presenceOfElementLocated(By.tagName("body"))
                ));
            } catch (Exception e) {
                System.out.println("Cycle " + cycle + ": Page load timeout, continuing...");
            }

            // Extended wait for AJAX content
            Thread.sleep(4000); // Increased to 4s for Zepto's slow AJAX

            // Collect product URLs
            List<WebElement> productLinks = driver.findElements(By.xpath("//a[contains(@href, '/pn/')]"));
            int currentLinkCount = productLinks.size();
            for (WebElement link : productLinks) {
                String href = link.getAttribute("href");
                if (href != null && !href.isEmpty() && href.contains("/pn/")) {
                    productUrlSet.add(href);
                    System.out.println("Cycle " + cycle + ": Added URL: " + href); // Log each URL
                } else {
                    System.out.println("Cycle " + cycle + ": Skipped invalid URL: " + href);
                }
            }

            int newSize = productUrlSet.size();
            long newHeight = (Long) ((JavascriptExecutor) driver).executeScript("return document.body.scrollHeight");

            System.out.println("Cycle " + cycle + ": Found " + newSize + " unique products (links on page: " +
                              currentLinkCount + "), height: " + newHeight + " (prev: " + previousHeight + ")");

            // Check for no change
            if (newSize == previousSize && newHeight == previousHeight) {
                noChangeCount++;
                if (noChangeCount >= maxNoChange) {
                    System.out.println("No new content for " + maxNoChange + " cycles. Stopping scroll.");
                    break;
                }
            } else {
                noChangeCount = 0;
            }

            previousSize = newSize;
            previousHeight = newHeight;

            Thread.sleep(1000);
        }

        List<String> productUrls = new ArrayList<>(productUrlSet);
        System.out.println("Total products found after scrolling: " + productUrls.size());
        if (productUrls.size() <= 4) {
            System.out.println("ERROR: Only " + productUrls.size() + " products found. Check logs for issues.");
        } else if (productUrls.size() <= 143) {
            System.out.println("WARNING: Still low count (" + productUrls.size() + "). Check logs for stabilization point.");
        }

        for (String productUrl : productUrls) {
            // Retry mechanism for page load
            boolean pageLoaded = false;
            int retries = 0;
            int maxRetries = 3;
            while (!pageLoaded && retries < maxRetries) {
                try {
                    driver.get(productUrl);
                    wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")));
                    pageLoaded = true;
                } catch (Exception e) {
                    retries++;
                    System.out.println("Failed to load product URL: " + productUrl + " (Retry " + retries + "/" + maxRetries + ")");
                    Thread.sleep(1000);
                }
            }
            if (!pageLoaded) {
                System.out.println("Skipping product URL after max retries: " + productUrl);
                continue;
            }

            // Extract name
            try {
                WebElement nameElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[@class='cp62rX c9OiKy c7CsPX']")));
                newName = nameElement.getText();
            } catch (NoSuchElementException e) {
                try {
                    WebElement nameElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class,'u-flex u-items-center')]//h1")));
                    newName = nameElement.getText();
                } catch (Exception ex) {
                    newName = "NA";
                }
            }
            System.out.println("Name: " + newName);
            System.out.println("headercount = " + headercount);
            headercount++;

            // Extract brand
            try {
                WebElement brandSection = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//p[@class='ccJTKY c9OiKy cL9VE0']")));
                brand = brandSection.getText().replace("brand", "").trim();
            } catch (Exception e) {
                if (newName != null && !newName.isEmpty()) {
                    String[] parts = newName.split(" ");
                    brand = (parts.length > 1) ? parts[0] + " " + parts[1] : parts[0];
                } else {
                    brand = "NA";
                }
            }
            System.out.println("Brand: " + brand);

            // Extract UOM
            try {
                WebElement uomEl = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@class='font-bold']")));
                uomValue = uomEl.getText().replace("Net Qty:", "").trim();
            } catch (Exception e) {
                uomValue = "NA";
            }
            System.out.println("UOM: " + uomValue);

            // Extract SP
            try {
                WebElement sp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@class='TvBIc']")));
                String originalSp = sp.getText().replace("₹", "");
                spValue = originalSp.trim();
            } catch (Exception e) {
                try {
                    WebElement sp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//p//span[@class='TvBIc']")));
                    String originalSp = sp.getText().replace("₹", "");
                    spValue = originalSp.trim();
                } catch (Exception ex) {
                    spValue = "NA";
                }
            }
            System.out.println("SP: " + spValue);

            // Extract MRP
            try {
                WebElement mrp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@class='_IKg9 dUCkZ']")));
                String originalMrp = mrp.getText().replace("₹", "");
                mrpValue = originalMrp.trim();
            } catch (NoSuchElementException e) {
                try {
                    WebElement mrp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//p//span[@class='_IKg9 dUCkZ']")));
                    String originalMrp = mrp.getText().replace("₹", "");
                    mrpValue = originalMrp.trim();
                } catch (Exception ex) {
                    try {
                        WebElement mrp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id=\"product-features-wrapper\"]/div[1]/div/div[3]/div[1]/div[2]/p/span[2]")));
                        String originalMrp = mrp.getText().replace("₹", "");
                        mrpValue = originalMrp.trim();
                    } catch (Exception exx) {
                        mrpValue = spValue;
                    }
                }
            }
            System.out.println("MRP: " + mrpValue);

            // Extract Offer
            try {
                WebElement offerEl = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//p[@class='__2zFOa']")));
                offerValue = offerEl.getText().trim();
            } catch (Exception e) {
                offerValue = "NA";
            }
            System.out.println("Offer: " + offerValue);

            // Check Availability
            int result = 1;
            try {
                String pageSource = driver.getPageSource().toLowerCase();
                if (pageSource.contains("notify me")) {
                    result = 0;
                }
            } catch (Exception e) {
                System.out.println("Error checking availability: " + e.getMessage());
            }
            NewAvailability1 = String.valueOf(result);
            System.out.println("Availability: " + NewAvailability1);

            System.out.println("===================================== product count :" + headercount + "=====================================================");

            // Write to Excel
            Row dataRow = sheet.createRow(rowNum++);
            dataRow.createCell(0).setCellValue(page);
            dataRow.createCell(1).setCellValue(brand);
            dataRow.createCell(2).setCellValue(headline);
            dataRow.createCell(3).setCellValue(productUrl);
            dataRow.createCell(4).setCellValue(newName);
            dataRow.createCell(5).setCellValue(mrpValue);
            dataRow.createCell(6).setCellValue(spValue);
            dataRow.createCell(7).setCellValue(uomValue);
            dataRow.createCell(8).setCellValue(NewAvailability1);
            dataRow.createCell(9).setCellValue(offerValue);

            System.out.println("---");

            Thread.sleep(500 + (int)(Math.random() * 500));
        }

        // Save the workbook
        try {
            new File(".\\Output").mkdirs();
            FileOutputStream out = new FileOutputStream(".\\Output\\Zepto_tissues_" + timestamp + ".xlsx");
            workbook.write(out);
            out.close();
            System.out.println("Excel file saved successfully.");
        } catch (Exception e) {
            System.out.println("Error saving Excel: " + e.getMessage());
        }

        driver.quit();
    }
}