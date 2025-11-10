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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/* -------------------------------------------------------------
   Simple POJO that replaces the missing AssortmentNew.dummy12.Product
   ------------------------------------------------------------- */
class Product {
    String name;
    String mrp;
    String sp;
    String uom;
    int    availability;
    String offer;
    String url;

    public Product(String name, String mrp, String sp, String uom,
                   int availability, String offer, String url) {
        this.name        = name;
        this.mrp         = mrp;
        this.sp          = sp;
        this.uom         = uom;
        this.availability= availability;
        this.offer       = offer;
        this.url         = url;
    }
}

/* -------------------------------------------------------------
   Main scraper class
   ------------------------------------------------------------- */
public class GroceryVeg {

    private static final List<Product> productList = new ArrayList<>();

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            // Hide webdriver flag
            js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

            driver.get("https://www.swiggy.com/instamart");

            // ---------- 1. Set location ----------
            setLocation(wait, "560001");

            // ---------- 2. Open Fresh Vegetables ----------
            wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//div[contains(text(),'Fresh Vegetables')]"))).click();
            Thread.sleep(3000);

            // ---------- 3. Get sub-categories ----------
            Thread.sleep(2000);
            List<WebElement> subcategories = wait.until(
                    ExpectedConditions.presenceOfAllElementsLocatedBy(
                            By.xpath("//div[@class='item-wrapper']")));

            System.out.println("Found " + subcategories.size() + " sub-categories.");

            for (int i = 0; i < subcategories.size(); i++) {
                try {
                    // Re-fetch to avoid StaleElementReferenceException
                    subcategories = driver.findElements(
                            By.xpath("//div[@class='item-wrapper']"));
                    WebElement subCat = subcategories.get(i);
                    String subCatName = subCat.getText().trim();
                    if (subCatName.isEmpty()) continue;

                    System.out.println("\n=== Sub-category: " + subCatName + " ===");

                    scrollToElement(driver, subCat);
                    wait.until(ExpectedConditions.elementToBeClickable(subCat));
                    subCat.click();
                    Thread.sleep(3000);

                    // Load **all** products (infinite scroll)
                    loadAllProducts(driver, wait);

                    // Product cards
                    List<WebElement> productCards = driver.findElements(
                            By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));

                    System.out.println("   Products found: " + productCards.size());

                    for (int j = 0; j < productCards.size(); j++) {
                        try {
                            // Re-fetch cards each iteration
                            productCards = driver.findElements(
                                    By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
                            WebElement card = productCards.get(j);

                            // Click via the image inside the card
                            WebElement img = card.findElement(
                                    By.xpath(".//img[contains(@class,'_16I1D') or contains(@alt,'')]"));
                            scrollToElement(driver, img);
                            js.executeScript("arguments[0].click();", img);
                            Thread.sleep(2000);

                            // Dismiss possible tooltip
                            dismissTooltip(js);

                            // ---------- Extract details ----------
                            String name = getText(driver, wait,
                                    ".//div[@class='sc-gEvEer gmdlfz _1iFYi'] | //h1", "NA");
                            String mrp  = getText(driver, wait,
                                    "//div[@data-testid='item-mrp-price'] |//div[@class='sc-gEvEer gKjjWC _2KTMQ'] ", "NA");
                            String sp   = getText(driver, wait,
                                    "//div[@data-testid='item-offer-price'] |//div[@class='sc-gEvEer iQcBUp _1bWTz'] ", mrp);
                            String uom  = getText(driver, wait,
                                    "(//div[@class='_30iun'] | //div[@class='sc-gEvEer ymEfJ _11EdJ'])[2]", "NA");
                            int    avail= getAvailability(driver, wait);
                            String offer= getText(driver, wait,
                                    "//div[@class='sc-gEvEer bsYAwc _1WaLo']", "NA");
                            String url  = driver.getCurrentUrl();

                            productList.add(new Product(name, mrp, sp, uom, avail, offer, url));

                            // Back to list
                            driver.navigate().back();
                            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                                    By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]")));
                        } catch (Exception e) {
                            System.out.println("   Product " + (j + 1) + " error: " + e.getMessage());
                            driver.navigate().back();
                            Thread.sleep(1000);
                        }
                    }

                    // Back to sub-category page
                    driver.navigate().back();
                    Thread.sleep(2000);
                } catch (Exception e) {
                    System.out.println("Sub-category " + (i + 1) + " error: " + e.getMessage());
                }
            }

            // ---------- Write Excel ----------
            writeToExcel(productList);
            System.out.println("\nScraping finished – " + productList.size() + " products saved.");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    /* ----------------------------------------------------------- */
    private static void setLocation(WebDriverWait wait, String pin) throws InterruptedException {
        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//div[@data-testid='search-location']"))).click();

        WebElement input = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.xpath("//input[contains(@placeholder,'Search for area')]")));
        input.sendKeys(pin);
        Thread.sleep(2000);

        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//div[contains(@class,'icon-location-marker')]"))).click();
        Thread.sleep(1000);

        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//button/span[contains(text(),'Confirm')]"))).click();
        Thread.sleep(3000);
    }

    private static void scrollToElement(WebDriver driver, WebElement el) {
        ((JavascriptExecutor) driver).executeScript(
                "arguments[0].scrollIntoView({block:'center'});", el);
        try { Thread.sleep(500); } catch (Exception ignored) {}
    }

    private static void loadAllProducts(WebDriver driver, WebDriverWait wait) throws InterruptedException {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        int last = 0;
        while (true) {
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
            Thread.sleep(3000);
            List<WebElement> prods = driver.findElements(
                    By.xpath("//div[contains(@class,'_3Rr1X')] | //div[contains(@class,'ProductCard')]"));
            if (prods.size() <= last) break;
            last = prods.size();
        }
    }

    private static void dismissTooltip(JavascriptExecutor js) {
        try {
            js.executeScript(
                "let el = document.querySelector('div[data-testid=\"re-check-address-tooltip\"] button, div[role=\"button\"]');" +
                "if (el) el.click();");
        } catch (Exception ignored) {}
    }

    private static String getText(WebDriver driver, WebDriverWait wait, String relXpath, String def) {
        try {
            WebElement el = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(relXpath)));
            String txt = el.getText().trim();
            return txt.isEmpty() ? def : txt;
        } catch (Exception e) {
            return def;
        }
    }

    private static int getAvailability(WebDriver driver, WebDriverWait wait) {
    	 String availability = "NA";
         try {
             driver.findElement(By.id("add-to-cart-button"));
             availability = "1";
         } catch (Exception e) {
             try {
                 driver.findElement(By.xpath("//span[contains(text(),'Currently unavailable')]"));
                 availability = "0";
             } catch (Exception ex) {
                 availability = "NA";
             }
         }
		return 0;
    }

    private static void writeToExcel(List<Product> products) {
        Workbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("Fresh Vegetables");

        // Header
        Row hdr = sh.createRow(0);
        String[] cols = {"Name","MRP","SP","UOM","Availability","Offer","URL"};
        for (int i = 0; i < cols.length; i++) hdr.createCell(i).setCellValue(cols[i]);

        // Data
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

        // Auto-size
        for (int i = 0; i < cols.length; i++) sh.autoSizeColumn(i);

        String file = "swiggy_vegetables_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";
        try (FileOutputStream fos = new FileOutputStream(file)) {
            wb.write(fos);
            System.out.println("Excel written: " + file);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try { wb.close(); } catch (IOException ignored) {}
        }
    }
}