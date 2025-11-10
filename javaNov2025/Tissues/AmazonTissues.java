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

public class AmazonTissues {

    public static void main(String[] args) throws InterruptedException, IOException {
        // Initialize Excel workbook
        Workbook resultsWorkbook = new XSSFWorkbook();
        Sheet resultsSheet = resultsWorkbook.createSheet("Results");
        Row headerRow = resultsSheet.createRow(0);

        headerRow.createCell(0).setCellValue("url");
        headerRow.createCell(1).setCellValue("newName");
        headerRow.createCell(2).setCellValue("mrpValue");
        headerRow.createCell(3).setCellValue("spValue");
        headerRow.createCell(4).setCellValue("Uom");
        headerRow.createCell(5).setCellValue("availability");
        headerRow.createCell(6).setCellValue("offerValue");
        headerRow.createCell(7).setCellValue("Brand");
        headerRow.createCell(8).setCellValue("Special_Feature");
        headerRow.createCell(9).setCellValue("NumberofItems");
        headerRow.createCell(10).setCellValue("SheetCount");
        headerRow.createCell(11).setCellValue("PlyRating");
        headerRow.createCell(12).setCellValue("Recommended");
        headerRow.createCell(13).setCellValue("Material");
        headerRow.createCell(14).setCellValue("NetQuantity");
        headerRow.createCell(15).setCellValue("Itemdimensions");
        headerRow.createCell(16).setCellValue("ItemFor");

        int rowIndex = 1;
        int HeaderCount = 1;

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--start-maximized");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(50));
        driver.get("https://www.amazon.in/s?k=tissue+paper+in+amazon&adgrpid=70848027495&ext_vrnc=hi&hvadid=381527342785&hvdev=c&hvlocphy=9303417&hvnetw=g&hvqmt=b&hvrand=16488695727427016471&hvtargid=kwd-1410160802546&hydadcr=9512_1953230&mcid=77eaaa2b5d8b3778bb2f18c88af728cb&qid=1760684374&xpid=e6I-_YA8VwKpi");
        Thread.sleep(3000);
        Set<String> productUrlSet = new HashSet<>();
        boolean hasMorePages = true;
        int maxPages = 20; // Safety limit to prevent infinite loop
        int currentPage = 1;

        while (hasMorePages && currentPage <= maxPages) {
            // Wait for search results to be present
            try {
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@data-component-type='s-search-result']")));
            } catch (Exception e) {
                System.out.println("Timeout waiting for search results on page " + currentPage);
            }

            // Collect product URLs
            List<WebElement> productLinks = driver.findElements(By.xpath("//a[@class='a-link-normal s-no-outline']"));
            for (WebElement link : productLinks) {
                String href = link.getAttribute("href");
                if (href != null && href.contains("/dp/") && !href.contains("#")) {
                    if (!href.startsWith("http")) {
                        href = "https://www.amazon.in" + href;
                    }
                    productUrlSet.add(href);
                }
            }

            System.out.println("Page " + currentPage + ": Found " + productLinks.size() + " products");

            // Check for next page
            try {
                WebElement nextButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(@class, 's-pagination-next')]")));
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", nextButton);
                currentPage++;
                Thread.sleep(2000); // Wait for page to load after click
            } catch (Exception e) {
                hasMorePages = false;
                System.out.println("No more pages or error navigating to next page.");
            }
        }

        System.out.println("Total unique product URLs found: " + productUrlSet.size());

        for (String prodUrl : productUrlSet) {
            driver.get(prodUrl);

            // Wait for page to load
            System.out.println("productcount" + HeaderCount++);
            try {
                wait.until(ExpectedConditions.jsReturnsValue("return document.readyState === 'complete';"));
            } catch (Exception e) {
                System.out.println("Page load timeout for " + prodUrl + ", continuing...");
            }

            // Click "Show more" button if present
            try {
                WebElement showMoreButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//span[contains(@class, 'a-expander-prompt')])[3]")));
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", showMoreButton);
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class, 'a-expander-content')]")));
                System.out.println("Clicked 'Show more' button");
            } catch (Exception e) {
                System.out.println("No 'Show more' button found or error clicking it");
            }

            // Extract product details
            String newName = "NA";
            try {
                WebElement nameElement = driver.findElement(By.xpath("//span[@id='productTitle']"));
                newName = nameElement.getText().trim();
                System.out.println("Name: " + newName);
            } catch (NoSuchElementException e) {
                System.out.println("Name: " + newName);
            }

            // MRP
            String mrpValue = "NA";
            try {
                WebElement mrp = driver.findElement(By.xpath("(//span[@class='a-price a-text-price'])[1]"));
                String originalMrp1 = mrp.getText();
                mrpValue = originalMrp1.replace("₹", "").replace("M.R.P.: ", "").replaceAll("[^0-9.]", "");
                System.out.println("MRP: " + mrpValue);
            } catch (NoSuchElementException e) {
                try {
                    WebElement bmrp = driver.findElement(By.xpath("(//span[@class='a-price a-text-price'])[1]//span"));
                    String MrpValue = bmrp.getText();
                    Pattern pattern = Pattern.compile("₹([\\d,]+\\.?\\d*)");
                    Matcher matcher = pattern.matcher(MrpValue);
                    if (matcher.find()) {
                        mrpValue = matcher.group(1).replace(",", "");
                    }
                    System.out.println("MRP: " + mrpValue);
                } catch (Exception ds) {
                    mrpValue = "NA";
                }
            }

            // Selling Price
            String spValue = "NA";
            try {
                WebElement sp = driver.findElement(By.xpath("//span[contains(@class,'a-price') and contains(@class, 'priceToPay')]"));
                String originalSp1 = sp.getText();
                Pattern pattern = Pattern.compile("₹([\\d,]+\\.?\\d*)");
                Matcher matcher = pattern.matcher(originalSp1);
                if (matcher.find()) {
                    spValue = matcher.group(1).replace(",", "");
                }
                System.out.println("SP: " + spValue);
            } catch (Exception e) {
                spValue = mrpValue;
                System.out.println("SP: " + spValue);
            }

            // Offer
            String offerValue = "NA";
            try {
                WebElement offer = driver.findElement(By.xpath("//span[contains(@class,'savingPriceOverride')]"));
                String newOffer = offer.getText();
                offerValue = newOffer.replace("OFF", "Off").replace("%", "% Off");
                System.out.println("Offer: " + offerValue);
            } catch (Exception e) {
                System.out.println("Offer: " + offerValue);
            }

            // Availability
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
            System.out.println("Availability: " + availability);

            // UOM extraction
            String Uom = "NA";
            try {
                driver.findElement(By.xpath("//span[@id='inline-twister-expanded-dimension-text-size_name']"));
                Uom = "1";
            } catch (Exception e) {
                try {
                    driver.findElement(By.xpath("(//span[@class='a-size-base a-color-base inline-twister-dim-title-value a-text-bold'])[2]"));
                    Uom = "0";
                } catch (Exception ex) {
                    Uom = "NA";
                }
            }
            System.out.println("UOM: " + Uom);

            // Brand
            String Brand = "";
            try {
                Thread.sleep(2000);
                WebElement Bran = driver.findElement(By.xpath("(//span[contains(text(),'Brand')]/following::td)[1]"));
                Brand = Bran.getText();
            } catch (Exception e) {
                try {
                    WebElement Bran = driver.findElement(By.xpath("(//span[contains(text(),'Brand')]/following::td)[1]//span"));
                    Brand = Bran.getText();
                } catch (Exception e1) {
                    Brand = "NA";
                }
            }
            System.out.println("Brand: " + Brand);

            // Special Feature
            String Special_Feature = "";
            try {
                Thread.sleep(2000);
                WebElement Special_Feat = driver.findElement(By.xpath("((//span[contains(text(),'Special Feature')])/following::td)[1]"));
                Special_Feature = Special_Feat.getText();
            } catch (Exception e) {
                try {
                    WebElement Special_Feat = driver.findElement(By.xpath("(//span[contains(text(),'Special Feature')]/following::td)[1]//span"));
                    Special_Feature = Special_Feat.getText();
                } catch (Exception e1) {
                    Special_Feature = "NA";
                }
            }
            System.out.println("Special Feature: " + Special_Feature);

            // Number of Items
            String NumberofItems = "";
            try {
                Thread.sleep(2000);
                WebElement NumberofIte = driver.findElement(By.xpath("((//span[contains(text(),'Number ')])/following::td)[1]"));
                NumberofItems = NumberofIte.getText();
            } catch (Exception e) {
                try {
                    WebElement NumberofIte = driver.findElement(By.xpath("(//span[contains(text(),'Number')]/following::td)[1]//span"));
                    NumberofItems = NumberofIte.getText();
                } catch (Exception e1) {
                    NumberofItems = "NA";
                }
            }
            System.out.println("Number of Items: " + NumberofItems);

            // Sheet Count
            String SheetCount = "";
            try {
                Thread.sleep(2000);
                WebElement SheetCo = driver.findElement(By.xpath("((//span[contains(text(),'Sheet Count')])/following::td)[1]"));
                SheetCount = SheetCo.getText();
            } catch (Exception e) {
                try {
                    WebElement SheetCo = driver.findElement(By.xpath("(//span[contains(text(),'Sheet Count')]/following::td)[1]//span"));
                    SheetCount = SheetCo.getText();
                } catch (Exception e1) {
                    SheetCount = "NA";
                }
            }
            System.out.println("Sheet Count: " + SheetCount);

            // Ply Rating
            String PlyRating = "";
            try {
                Thread.sleep(2000);
                WebElement PlyRati = driver.findElement(By.xpath("((//span[contains(text(),'Ply Rating')])/following::td)[1]"));
                PlyRating = PlyRati.getText();
            } catch (Exception e) {
                try {
                    WebElement PlyRati = driver.findElement(By.xpath("(//span[contains(text(),'Ply Rating')]/following::td)[1]//span"));
                    PlyRating = PlyRati.getText();
                } catch (Exception e1) {
                    PlyRating = "NA";
                }
            }
            System.out.println("Ply Rating: " + PlyRating);

            // Recommended
            String Recommended = "";
            try {
                Thread.sleep(2000);
                WebElement Recommend = driver.findElement(By.xpath("((//span[contains(text(),'Recommended')])/following::td)[1]"));
                Recommended = Recommend.getText();
            } catch (Exception e) {
                try {
                    WebElement Recommend = driver.findElement(By.xpath("(//span[contains(text(),'Recommended')]/following::td)[1]//span"));
                    Recommended = Recommend.getText();
                } catch (Exception e1) {
                    Recommended = "NA";
                }
            }
            System.out.println("Recommended: " + Recommended);

            // Material
            String Material = "";
            try {
                Thread.sleep(2000);
                WebElement Materia = driver.findElement(By.xpath("((//span[contains(text(),'Material')])/following::td)[1]"));
                Material = Materia.getText();
            } catch (Exception e) {
                try {
                    WebElement Materia = driver.findElement(By.xpath("(//span[contains(text(),'Material')]/following::td)[1]//span"));
                    Material = Materia.getText();
                } catch (Exception e1) {
                    Material = "NA";
                }
            }
            System.out.println("Material: " + Material);

            // Net Quantity
            String NetQuantity = "";
            try {
                Thread.sleep(2000);
                WebElement NetQuant = driver.findElement(By.xpath("((//span[contains(text(),'Net Quantity')])/following::td)[1]"));
                NetQuantity = NetQuant.getText();
            } catch (Exception e) {
                try {
                    WebElement NetQuant = driver.findElement(By.xpath("(//span[contains(text(),'Net Quantity')]/following::td)[1]//span"));
                    NetQuantity = NetQuant.getText();
                } catch (Exception e1) {
                    NetQuantity = "NA";
                }
            }
            System.out.println("Net Quantity: " + NetQuantity);

            // Item dimensions
            String Itemdimensions = "";
            try {
                Thread.sleep(2000);
                WebElement Itemdimensio = driver.findElement(By.xpath("((//span[contains(text(),'Item dimensions')])/following::td)[1]"));
                Itemdimensions = Itemdimensio.getText();
            } catch (Exception e) {
                try {
                    WebElement Itemdimensio = driver.findElement(By.xpath("(//span[contains(text(),'Item dimensions')]/following::td)[1]//span"));
                    Itemdimensions = Itemdimensio.getText();
                } catch (Exception e1) {
                    Itemdimensions = "NA";
                }
            }
            System.out.println("Item dimensions: " + Itemdimensions);

            // Item Form
            String ItemFor = "";
            try {
                Thread.sleep(2000);
                WebElement ItemF = driver.findElement(By.xpath("((//span[contains(text(),'Item Form')])/following::td)[1]"));
                ItemFor = ItemF.getText();
            } catch (Exception e) {
                try {
                    WebElement ItemF = driver.findElement(By.xpath("(//span[contains(text(),'Item Form')]/following::td)[1]//span"));
                    ItemFor = ItemF.getText();
                } catch (Exception e1) {
                    ItemFor = "NA";
                }
            }
            System.out.println("Item Form: " + ItemFor);

            String url = prodUrl;

            // Write to Excel
            Row resultRow = resultsSheet.createRow(rowIndex++);
            resultRow.createCell(0).setCellValue(url);
            resultRow.createCell(1).setCellValue(newName);
            resultRow.createCell(2).setCellValue(mrpValue);
            resultRow.createCell(3).setCellValue(spValue);
            resultRow.createCell(4).setCellValue(Uom);
            resultRow.createCell(5).setCellValue(availability);
            resultRow.createCell(6).setCellValue(offerValue);
            resultRow.createCell(7).setCellValue(Brand);
            resultRow.createCell(8).setCellValue(Special_Feature);
            resultRow.createCell(9).setCellValue(NumberofItems);
            resultRow.createCell(10).setCellValue(SheetCount);
            resultRow.createCell(11).setCellValue(PlyRating);
            resultRow.createCell(12).setCellValue(Recommended);
            resultRow.createCell(13).setCellValue(Material);
            resultRow.createCell(14).setCellValue(NetQuantity);
            resultRow.createCell(15).setCellValue(Itemdimensions);
            resultRow.createCell(16).setCellValue(ItemFor);

            Thread.sleep(2000);
        }

        // Save Excel file
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        String outputFilePath = ".\\Output\\Amazon_Tissues & Disposables_" + timestamp + ".xlsx";

        // Write results to Excel file
        FileOutputStream outFile = new FileOutputStream(outputFilePath);
        resultsWorkbook.write(outFile);
        outFile.close();
        resultsWorkbook.close();
        driver.quit();
        System.out.println("Data written to " + outputFilePath);
    }
}