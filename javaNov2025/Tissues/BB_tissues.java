package AssortmentTissues;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

public class BB_tissues {

    public static void main(String[] args) throws InterruptedException, IOException {
        // Initialize Excel workbook
        Workbook resultsWorkbook = new XSSFWorkbook();
        Sheet resultsSheet = resultsWorkbook.createSheet("Results");
        Row headerRow = resultsSheet.createRow(0);
       
        headerRow.createCell(0).setCellValue("URL");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("MRP");
        headerRow.createCell(3).setCellValue("SP");
        headerRow.createCell(4).setCellValue("UOM");
        headerRow.createCell(5).setCellValue("Multiplier");
        headerRow.createCell(6).setCellValue("Availability");
        headerRow.createCell(7).setCellValue("Offer");

        int rowIndex = 1;
        int HeaderCount= 1;

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--start-maximized");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(50));
        driver.get("https://www.bigbasket.com/");
        Thread.sleep(3000);
        driver.get("https://www.bigbasket.com/pc/cleaning-household/disposables-garbage-bag/kitchen-rolls/?nc=ct-fa&sid=X_8h342ibWQBoWMOqHNrdV9saXN0kKJuZsOiY2OoNDE5fDE2MTOpYmF0Y2hfaWR4AKJhb8KidXLComFww6JsdM0MX6FvqnBvcHVsYXJpdHmlc3JfaWQBo21yaRw%3D");

        Set<String> productUrlSet = new HashSet<>();
        int maxCycles = 2; // Limit scrolling to 3 times
        int cycle = 0;
        long lastHeight = (Long) ((JavascriptExecutor) driver).executeScript("return document.body.scrollHeight");

        System.out.println("Starting scroll cycle. Initial height: " + lastHeight);

        while (cycle < maxCycles) {
            cycle++;

            // Scroll incrementally to bottom slowly
            long currentHeight = (Long) ((JavascriptExecutor) driver).executeScript("return document.body.scrollHeight");
            long pos = 0;
            long increment = 100;
            while (pos < currentHeight) {
                ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, " + pos + ");");
                pos += increment;
                Thread.sleep(400);
                try {
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//a[contains(@href, '/pd/')]")));
                } catch (Exception e) {
                    System.out.println("Cycle " + cycle + ": No new links during scroll, continuing...");
                }
            }

            // Wait for content to load
            try {
                wait.until(ExpectedConditions.or(
                    ExpectedConditions.jsReturnsValue("return document.readyState === 'complete';"),
                    ExpectedConditions.presenceOfElementLocated(By.tagName("body"))
                ));
            } catch (Exception e) {
                System.out.println("Cycle " + cycle + ": Page load timeout, continuing...");
            }

            // Collect product URLs
            List<WebElement> links = driver.findElements(By.xpath("//a[contains(@href, '/pd/')]"));
            for (WebElement link : links) {
                String href = link.getAttribute("href");
                if (href != null && href.contains("/pd/") && !href.contains("#")) {
                    productUrlSet.add(href);
                }
            }

            // Scroll back to top to trigger lazy loads if needed
            ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, 0);");
            Thread.sleep(500);
        }

        System.out.println("Total unique product URLs found: " + productUrlSet.size());

        for (String prodUrl : productUrlSet) {
            driver.get(prodUrl);

            // Wait for page to load
            System.out.println("-------------------------productcount " + (HeaderCount++) + "-------------------------");
            try {
                wait.until(ExpectedConditions.jsReturnsValue("return document.readyState === 'complete';"));
            } catch (Exception e) {
                System.out.println("Page load timeout for " + prodUrl + ", continuing...");
            }

            // Extract product details
            String newName = "NA";
            try {
                WebElement nameElement = driver.findElement(By.xpath("//h1[@class='Description___StyledH-sc-82a36a-2 bofYPK']"));
                newName = nameElement.getText();
                System.out.println("Name: " + newName);
            } catch (NoSuchElementException e) {
                try {
                    WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/section[1]/div[2]/section[1]/h1"));
                    newName = nameElement.getText();
                    System.out.println("Name: " + newName);
                } catch (NoSuchElementException ex) {
                    System.out.println("Name: " + newName);
                }
            }

            // MRP
            String mrpValue = "NA";
            try {
                WebElement mrp = driver.findElement(By.xpath("//td[@class='line-through p-0']"));
                String originalMrp1 = mrp.getText();
                mrpValue = originalMrp1.replace("₹", "").replace("MRP: ", "");
                System.out.println("MRP: " + mrpValue);
            } catch (NoSuchElementException e) {
                try {
                    WebElement bmrp = driver.findElement(By.xpath("//td[@class='Description___StyledTd-sc-82a36a-4 fLZywG']"));
                    String MrpValue = bmrp.getText();
                    Pattern pattern = Pattern.compile("₹(\\d+\\.?\\d*)");
                    Matcher matcher = pattern.matcher(MrpValue);
                    if (matcher.find()) {
                        mrpValue = matcher.group(1);
                    }
                    System.out.println("MRP: " + mrpValue);
                } catch (Exception ds) {
                    System.out.println("MRP: " + mrpValue);
                }
            }

            // Selling Price
            String spValue = "NA";
            try {
                WebElement sp = driver.findElement(By.xpath("//td[@class='Description___StyledTd-sc-82a36a-4 fLZywG']"));
                String originalSp1 = sp.getText();
                Pattern pattern = Pattern.compile("₹(\\d+\\.?\\d*)");
                Matcher matcher = pattern.matcher(originalSp1);
                if (matcher.find()) {
                    spValue = matcher.group(1);
                }
                System.out.println("SP: " + spValue);
            } catch (Exception e) {
                spValue = mrpValue;
                System.out.println("SP: " + spValue);
            }

            // Offer
            String offerValue = "NA";
            try {
                WebElement offer = driver.findElement(By.xpath("//*[@id='siteLayout']/div/div/section[1]/div[2]/section[1]/table/tr[3]/td[2]"));
                String newOffer = offer.getText();
                offerValue = newOffer.replace("OFF", "Off");
                System.out.println("Offer: " + offerValue);
            } catch (Exception e) {
                System.out.println("Offer: " + offerValue);
            }

            // Availability
            String availability = "NA";
            try {
                driver.findElement(By.xpath("(//button[text()='Add to basket'])[1]"));
                availability = "1";
            } catch (Exception e) {
                try {
                    driver.findElement(By.xpath("(//button[text()='Notify Me'])[1]"));
                    availability = "0";
                } catch (Exception ex) {
                    availability = "NA";
                }
            }
            System.out.println("Availability: " + availability);

            // UOM extraction from product name
            String uom = "NA";
            try {
                // Regex to match patterns like "1 pc (100 Pulls)", "22 X 22 cm", "100 Serviettes", "1 Ply", "3 x 100 pcs"
                Pattern uomPattern = Pattern.compile("(\\d+\\s*(?:pc|Pulls|Serviettes|Ply|pcs|Pack)(?:\\s*\\(\\d+\\s*(?:Pulls|Serviettes)\\))?|\\d+\\s*x\\s*\\d+\\s*cm)", Pattern.CASE_INSENSITIVE);
                Matcher uomMatcher = uomPattern.matcher(newName);
                if (uomMatcher.find()) {
                    uom = uomMatcher.group(1);
                }
            } catch (Exception e) {
                uom = "NA";
            }
            System.out.println("UOM: " + uom);

            

            String url = prodUrl;

            // Write to Excel
            Row resultRow = resultsSheet.createRow(rowIndex++);
            resultRow.createCell(0).setCellValue(url);
            resultRow.createCell(1).setCellValue(newName);
            resultRow.createCell(2).setCellValue(mrpValue);
            resultRow.createCell(3).setCellValue(spValue);
            resultRow.createCell(4).setCellValue(uom);
            resultRow.createCell(5).setCellValue(availability);
            resultRow.createCell(6).setCellValue(offerValue);
            

            Thread.sleep(2000);
        }
        
        // Save Excel file
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        String outputFilePath = ".\\Output\\BB_Kitchen Rolls _" + timestamp + ".xlsx";

        // Write results to Excel file
        FileOutputStream outFile = new FileOutputStream(outputFilePath);
        resultsWorkbook.write(outFile);
        outFile.close();
        resultsWorkbook.close();
        driver.quit();
        System.out.println("Data written to " + outputFilePath);
    }
}