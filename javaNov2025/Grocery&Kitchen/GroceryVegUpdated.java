package freshVeg;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/* -------------------------------------------------------------
   POJO – unchanged
   ------------------------------------------------------------- */
class Product {
    String name, mrp, sp, uom, offer, url;
    int    availability;

    public Product(String name, String mrp, String sp, String uom,
                   int availability, String offer, String url) {
        this.name = name; this.mrp = mrp; this.sp = sp; this.uom = uom;
        this.availability = availability; this.offer = offer; this.url = url;
    }
}

/* -------------------------------------------------------------
   Main scraper – **optimised**
   ------------------------------------------------------------- */
public class GroceryVeg {

    private static final List<Product> productList = new ArrayList<>();

    public static void main(String[] args) {
        // ---------- Chrome options – fast & headless-friendly ----------
    	ChromeOptions options = new ChromeOptions();
	        	options.addArguments("--incognito"); // Run Chrome in headless mode
		    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");
        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        // ---------- Explicit wait – reuse one instance ----------
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            js.executeScript("Object.defineProperty(navigator,'webdriver',{get:()=>undefined});");
            driver.get("https://www.swiggy.com/instamart");

            setLocation(wait, "560001");

            wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//div[contains(text(),'Fresh Vegetables')]"))).click();
                                     // minimal pause
            Thread.sleep(5000);
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
                    subCat.click();                         // no extra wait – scroll already loaded it
                   

                    loadAllProducts(driver, wait);          // fast infinite-scroll

                    List<WebElement> productCards = driver.findElements(
                            By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
                    System.out.println("   Products: " + productCards.size());

                    for (int j = 0; j < productCards.size(); j++) {
                        // ----- retry block for occasional error pages -----
                        if (!handleErrorPage(wait, driver)) break;   // abort subcategory if too many retries

                        productCards = driver.findElements(
                                By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
                        WebElement card = productCards.get(j);

                        WebElement img = card.findElement(
                                By.xpath(".//img[contains(@class,'_16I1D') or contains(@alt,'')]"));
                        scrollToElement(driver, img);
                        js.executeScript("arguments[0].click();", img);
                                           // reduced

                        dismissTooltip(js);

                        // ----- Extract with **single** wait per field -----
                        String name = getText(wait, "//div[@class='sc-gEvEer gmdlfz _1iFYi'] | //h1", "NA");
                        String mrp  = getText(wait, "//div[@data-testid='item-mrp-price'] | //div[@class='sc-gEvEer gKjjWC _2KTMQ']", "NA");
                        String sp   = getText(wait, "//div[@data-testid='item-offer-price'] | //div[@class='sc-gEvEer iQcBUp _1bWTz']", mrp);
                        String uom  = getText(wait, "(//div[@class='_30iun'] | //div[@class='sc-gEvEer ymEfJ _11EdJ'])[2]", "NA");
                        int    avail= getAvailability(driver);
                        String offer= getText(wait, "//div[@class='sc-gEvEer bsYAwc _1WaLo']", "NA");
                        String url  = driver.getCurrentUrl();
                        System.out.println(name + " | " + mrp + " | " + sp + " | " + uom + " | " + avail + " | " + offer + " | " + url);

                        productList.add(new Product(name, mrp, sp, uom, avail, offer, url));

                        driver.navigate().back();
                        wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                                By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]")));
                    }

                    driver.navigate().back();               // back to sub-cat list
                    
                } catch (Exception e) {
                    System.out.println("Sub-cat " + (i + 1) + " error: " + e.getMessage());
                }
            }

            writeToExcel(productList);
            System.out.println("\nFinished – " + productList.size() + " items saved.");

        } catch (Exception e) { e.printStackTrace(); }
        finally { driver.quit(); }
    }

    /* ----------------------------------------------------------- */
    private static void setLocation(WebDriverWait wait, String pin) throws InterruptedException {
        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//div[@data-testid='search-location']"))).click();

        WebElement input = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.xpath("//input[contains(@placeholder,'Search for area')]")));
        input.sendKeys(pin);
        

        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//div[contains(@class,'icon-location-marker')]"))).click();
        

        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//button/span[contains(text(),'Confirm')]"))).click();
        
    }

    private static void scrollToElement(WebDriver driver, WebElement el) {
        ((JavascriptExecutor) driver).executeScript(
                "arguments[0].scrollIntoView({block:'center'});", el);
    }

    private static void loadAllProducts(WebDriver driver, WebDriverWait wait) throws InterruptedException {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        int last = 0;
        while (true) {
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
                                    // faster scroll
            List<WebElement> prods = driver.findElements(
                    By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
            if (prods.size() <= last) break;
            last = prods.size();
        }
    }

    private static void dismissTooltip(JavascriptExecutor js) {
        try {
            js.executeScript(
                "let el=document.querySelector('div[data-testid=\"re-check-address-tooltip\"] button, div[role=\"button\"]');"
              + "if(el)el.click();");
        } catch (Exception ignored) {}
    }

    /** One wait per field – avoids repeated driver.findElement */
    private static String getText(WebDriverWait wait, String xpath, String def) {
        try {
            return wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)))
                       .getText().trim();
        } catch (Exception e) { return def; }
    }

    private static int getAvailability(WebDriver driver) {
        try { driver.findElement(By.id("add-to-cart-button")); return 1; }
        catch (Exception e) {
            try { driver.findElement(By.xpath("//span[contains(text(),'Currently unavailable')]")); return 0; }
            catch (Exception ex) { return 999; }
        }
    }

    /** Handles "Something went wrong" page – returns false on max retries 
     * @throws InterruptedException */
    private static boolean handleErrorPage(WebDriverWait wait, WebDriver driver) throws InterruptedException {
        int maxRetries = 4;
        int attempt = 0;

        while (attempt < maxRetries) {
            try {
                // Check if "Something went wrong" or "Try Again" exists
                List<WebElement> errorElements = driver.findElements(
                    By.xpath("//*[contains(text(),'Something went wrong')] | //button[contains(.,'Try Again')]")
                );

                if (!errorElements.isEmpty()) {
                    System.out.println("Error page detected – attempt " + (attempt + 1));

                    try {
                        // Wait for the "Try Again" button to be clickable
                        WebElement tryAgainBtn = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[contains(.,'Try Again')]")
                        ));

                        // Click using JavaScript (more reliable)
                        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", tryAgainBtn);
                        System.out.println("Clicked 'Try Again' button.");

                        // Wait briefly for page reload
                        Thread.sleep(3000);

                    } catch (Exception e) {
                        System.out.println("Try-Again button not clickable – forcing refresh...");
                        driver.navigate().refresh();
                        Thread.sleep(3000);
                    }

                    // After handling, recheck if still on error page
                    attempt++;
                    continue;
                }

                // No error message found — exit successfully
                return true;

            } catch (Exception e) {
                System.out.println("Unexpected error while handling error page: " + e.getMessage());
                driver.navigate().refresh();
                Thread.sleep(3000);
            }

            attempt++;
        }

        System.out.println("Max retries exhausted – skipping this subcategory.");
        return false;
    }

    /* ----------------------------------------------------------- */
    private static void writeToExcel(List<Product> products) {
        Workbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("Fresh Vegetables");
        Row hdr = sh.createRow(0);
        String[] cols = {"Name","MRP","SP","UOM","Availability","Offer","URL"};
        for (int i = 0; i < cols.length; i++) hdr.createCell(i).setCellValue(cols[i]);

        int r = 1;
        for (Product p : products) {
            Row row = sh.createRow(r++);
            row.createCell(0).setCellValue(p.name);
            row.createCell(1).setCellValue(p.mrp);
            row.createCell(2).setCellValue(p.sp);
            row.createCell(3).setCellValue(p.uom);
            row.createCell(4).setCellValue(p.availability);
            row.createCell(5).setCellValue(p.offer);
            row.createCell(6).setCellValue(p.url);
        }
        for (int i = 0; i < cols.length; i++) sh.autoSizeColumn(i);

        String file = "swiggy_vegetables_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";
        try (FileOutputStream fos = new FileOutputStream(file)) {
            wb.write(fos);
            System.out.println("Excel written: " + file);
        } catch (IOException e) { e.printStackTrace(); }
        finally { try { wb.close(); } catch (IOException ignored) {} }
    }
}