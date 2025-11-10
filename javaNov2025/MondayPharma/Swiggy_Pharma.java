package Swigyy_pharma;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Swiggy_Pharma {

    public static void main(String[] args) throws Exception {
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--disable-blink-features=AutomationControlled");
        options.addArguments("--disable-extensions");
        options.addArguments("--start-maximized");
        options.addArguments("--disable-gpu");
        options.addArguments("--remote-debugging-port=9222");
        options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));
        options.setExperimentalOption("useAutomationExtension", false);
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36");

        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.managed_default_content_settings.images", 2);
        prefs.put("profile.managed_default_content_settings.stylesheets", 2);
        prefs.put("profile.managed_default_content_settings.fonts", 2);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        // CRITICAL: Hide webdriver flag
        ((JavascriptExecutor) driver).executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => false});");

        int count = 0;
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
            String filePath = ".\\input-data\\Phrma input data-1.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Swiggy");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                    InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
                    UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), Availability = new ArrayList<>(),
                    Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();

            // Extract URLs from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);

                if (i == 0) {
                    continue;
                }

                Cell inputPidCell = row.getCell(0);
                Cell inputCityCell = row.getCell(1);
                Cell inputNameCell = row.getCell(2);
                Cell inputSizeCell = row.getCell(3);
                Cell newProductCodeCell = row.getCell(4);
                Cell urlCell = row.getCell(5);
                Cell uomCell = row.getCell(6);
                Cell multiplierCell = row.getCell(7);
                Cell availabilityCell = row.getCell(8);
                Cell pinCodeCell = row.getCell(9);
                Cell oldNameCell = row.getCell(10);

                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING)
                            ? inputPidCell.getStringCellValue()
                            : "";
                    String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING)
                            ? inputCityCell.getStringCellValue()
                            : "";
                    String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING)
                            ? inputNameCell.getStringCellValue()
                            : "";
                    String size = (inputSizeCell != null && inputSizeCell.getCellType() == CellType.STRING)
                            ? inputSizeCell.getStringCellValue()
                            : "";
                    String productCode = (newProductCodeCell != null && newProductCodeCell.getCellType() == CellType.STRING)
                            ? newProductCodeCell.getStringCellValue()
                            : "";
                    String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING)
                            ? uomCell.getStringCellValue()
                            : "";
                    String mulitiplier = (multiplierCell != null && multiplierCell.getCellType() == CellType.STRING)
                            ? multiplierCell.getStringCellValue()
                            : "";
                    String availability = (availabilityCell != null && availabilityCell.getCellType() == CellType.STRING)
                            ? availabilityCell.getStringCellValue()
                            : "";
                    String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING)
                            ? pinCodeCell.getStringCellValue()
                            : "";
                    String namecheck = (oldNameCell != null && oldNameCell.getCellType() == CellType.STRING)
                            ? oldNameCell.getStringCellValue()
                            : "";

                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
                    UOM.add(uom);
                    Mulitiplier.add(mulitiplier);
                    Availability.add(availability);
                    Pincode.add(locationSet);
                    NameForCheck.add(namecheck);
                }
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
            headerRow.createCell(12).setCellValue("Offer");

            int rowIndex = 1;
            int headercount = 1;
            String currentPin = null;

            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String locationSet = Pincode.get(i);

                try {
                    if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
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
                        continue;
                    }

                    if (currentPin == null || !currentPin.equals(locationSet)) {
                        driver.get("https://www.swiggy.com/");
                        driver.manage().window().maximize();
                        Thread.sleep(3000);

                        try {
                            WebElement location = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@class='_5ZhdF _3GoNS _1LZf8']")));
                            location.click();
                            Thread.sleep(1000);
                            WebElement locationField = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@id='location']")));
                            wait.until(ExpectedConditions.elementToBeClickable(locationField));
                            locationField.clear();
                            locationField.clear();
                            locationField.clear();
                            Thread.sleep(1000);
                            System.out.println("Sending pin code: " + locationSet);
                            locationField.sendKeys(locationSet);
                            Thread.sleep(1000);
                            wait.until(ExpectedConditions.textToBePresentInElementValue(locationField, locationSet));
                            WebElement suggestion = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//div[@class='_2NKIb']//div[@class='kuQWc'])[1]")));
                            suggestion.click();

                            currentPin = locationSet;
                            Thread.sleep(3000);
                        } catch (Exception e) {
                            System.out.println("Pincode setup failed for: " + locationSet);
                        }
                    }

                    driver.manage().window().maximize();
                    Thread.sleep(2000 + (long)(Math.random() * 3000)); // Human delay

                    driver.get(url);

                    // Human-like scroll
                    ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight / 3);");
                    Thread.sleep(800);
                    ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, 0);");
                    Thread.sleep(500);

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
                            resultRow.createCell(12).setCellValue(offerValue);
                            System.out.println("Something went wrong found, URL skipped: " + url);
                            continue;
                        }
                    } catch (NoSuchElementException e) {
                        // Page loaded, proceed
                    }

                    try {
                        WebElement nameElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1")));
                        newName = nameElement.getText();
                        System.out.println("Name: " + newName);
                    } catch (Exception ed) {
                        try {
                            WebElement nameElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='sc-aXZVg aUBRn _1iFYi']")));
                            newName = nameElement.getText();
                            System.out.println(newName);
                        } catch (Exception e1) {
                            newName = "NA";
                        }
                    }

                    System.out.println("================headercount :" + headercount);
                    headercount++;
                    System.out.println("======================================+ " + (headercount++) + " +==================================================================");

                    // SP (Selling Price)
                    try {
                        int retries = 3;
                        for (int attempt = 0; attempt < retries; attempt++) {
                            try {
                                WebElement sp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@data-testid='item-offer-price']")));
                                originalSp1 = sp.getText().replace("₹", "").trim();
                                if (!originalSp1.isEmpty() && !originalSp1.equalsIgnoreCase("NaN")) {
                                    break;
                                }
                                Thread.sleep(1000);
                            } catch (Exception ex) {
                                System.out.println("Retry " + (attempt + 1) + " for SP");
                            }
                        }

                        if (originalSp1.isEmpty() || originalSp1.equalsIgnoreCase("NaN")) {
                            originalSp1 = (String) ((JavascriptExecutor) driver).executeScript(
                                    "return document.querySelector('[data-testid=\"item-offer-price\"]')?.textContent || '';"
                            );
                            originalSp1 = originalSp1.replace("₹", "").trim();
                        }

                        spValue = originalSp1.isEmpty() ? "NA" : originalSp1;
                        System.out.println("SP: " + spValue);
                    } catch (Exception ee) {
                        spValue = "NA";
                        System.out.println("Failed to extract SP: ");
                    }

                    // MRP (Maximum Retail Price)
                    try {
                        int retries = 3;
                        for (int attempt = 0; attempt < retries; attempt++) {
                            try {
                                WebElement mrp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@data-testid='item-mrp-price']")));
                                originalMrp1 = mrp.getText().replace("₹", "").trim();
                                if (!originalMrp1.isEmpty() && !originalMrp1.equalsIgnoreCase("NaN")) {
                                    break;
                                }
                                Thread.sleep(1000);
                            } catch (Exception ex) {
                                System.out.println("Retry " + (attempt + 1) + " for MRP");
                            }
                        }

                        if (originalMrp1.isEmpty() || originalMrp1.equalsIgnoreCase("NaN")) {
                            originalMrp1 = (String) ((JavascriptExecutor) driver).executeScript(
                                    "return document.querySelector('[data-testid=\"item-mrp-price\"]')?.textContent || '';"
                            );
                            originalMrp1 = originalMrp1.replace("₹", "").trim();
                        }

                        mrpValue = originalMrp1.isEmpty() ? spValue : originalMrp1;
                        System.out.println("MRP: " + mrpValue);
                    } catch (NoSuchElementException eg) {
                        try {
                            WebElement mrp = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]"));
                            originalMrp1 = mrp.getText().replace("₹", "").trim();
                            mrpValue = originalMrp1.isEmpty() ? spValue : originalMrp1;
                            System.out.println("MRP (alternative): " + mrpValue);
                        } catch (Exception ex) {
                            try {
                                originalMrp1 = (String) ((JavascriptExecutor) driver).executeScript(
                                        "return document.querySelector('#product-details-page-container div div div div div div div div div div div:nth-child(2)')?.textContent || '';"
                                );
                                originalMrp1 = originalMrp1.replace("₹", "").trim();
                                mrpValue = originalMrp1.isEmpty() ? spValue : originalMrp1;
                                System.out.println("MRP (JS fallback): " + mrpValue);
                            } catch (Exception ex2) {
                                mrpValue = spValue;
                                System.out.println("Failed to extract MRP, defaulting to SP: " + mrpValue);
                            }
                        }
                    }

                    // UOM
                    try {
                        WebElement webUom1 = driver.findElement(By.xpath("//div[@class='sc-gEvEer ymEfJ _11EdJ']"));
                        String webUom2 = webUom1.getText();
                        webUom = webUom2;
                        System.out.println("UOM: " + webUom);
                    } catch (Exception u) {
                        try {
                            WebElement webUom1 = driver.findElement(By.xpath("//div[@class='sc-eqUAAy dEjugH _1TwvP']"));
                            String webUom2 = webUom1.getText();
                            webUom = webUom2;
                            System.out.println("UOM (alternative): " + webUom);
                        } catch (Exception ex) {
                            webUom = "NA";
                            System.out.println("Failed to extract UOM: " + ex.getMessage());
                        }
                    }

                    // Availability
                    int result = 1;
                    if (url.contains("NA")) {
                        NewAvailability1 = "NA";
                    } else {
                        try {
                            String[] textsToCheck = {
                                    "Currently Unavailable",
                                    "Currently out of stock in this area.",
                                    "Sold Out",
                                    "Unavailable"
                            };
                            String pageSource = driver.getPageSource();
                            boolean isTextPresent = false;

                            for (String text : textsToCheck) {
                                if (pageSource.contains(text)) {
                                    isTextPresent = true;
                                    break;
                                }
                            }

                            result = isTextPresent ? 0 : 1;
                            System.out.println("Availability result: " + result);
                        } catch (Exception ehf) {
                            System.out.println("Error checking availability: " + ehf.getMessage());
                            result = -1;
                        }
                    }
                    NewAvailability1 = String.valueOf(result);

                    // Multiplier
                    multiplier = calculateMultiplier(size, webUom);
                    System.out.println("Multiplier: " + multiplier);

                    // Offer
                    if (mrpValue.equals(spValue)) {
                        offerValue = "NA";
                    } else {
                        try {
                            WebElement Offer = driver.findElement(By.xpath("(//div[@data-testid='item-offer-label-discount-text'])[1]"));
                            offerValue = Offer.getText();
                            System.out.println("Offer: " + offerValue);
                        } catch (NoSuchElementException ab) {
                            try {
                                WebElement Offer = driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/div[2]/div[1]/div/div/div[1]/div[2]/div"));
                                offerValue = Offer.getText();
                                System.out.println("Offer (alternative): " + offerValue);
                            } catch (Exception S) {
                                offerValue = "NA";
                                System.out.println("No offer found, defaulting to NA");
                            }
                        }
                    }

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
                    resultRow.createCell(12).setCellValue(offerValue);

                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
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
                    resultRow.createCell(12).setCellValue(offerValue);
                    System.out.println("Failed to extract data for URL: " + url);
                }
            }

            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = ".\\Output\\Swiggy_Pharma_" + timestamp + ".xlsx";

            FileOutputStream outFile = new FileOutputStream(outputFilePath);
            resultsWorkbook.write(outFile);
            outFile.close();

            System.out.println("Output file saved: " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
                System.out.println("Scraping Completed.");
            }
        }
    }

    private static double calculateMultiplier(String inputUom, String outputUom) {
        try {
            String input = inputUom.toLowerCase().trim();
            String output = outputUom.toLowerCase().trim();

            if (output.matches("\\d+\\s*pack.*")) {
                int inputPackCount = extractPackCount(input);
                int outputPackCount = extractPackCount(output);

                if (inputPackCount > 0 && outputPackCount > 0) {
                    if (inputPackCount == outputPackCount) {
                        return 1.0;
                    }
                }
            }

            double inputTotal = calculateTotalFromUom(input);
            double outputTotal = calculateTotalFromUom(output);

            if (outputTotal == 0) {
                return 0;
            }

            if (output.contains("pack") && !output.matches(".*\\(.*\\).*")) {
                if (input.contains("pack")) {
                    int inputPackCount = extractPackCount(input);
                    int outputPackCount = extractPackCount(output);
                    if (inputPackCount > 0 && outputPackCount > 0 && inputPackCount == outputPackCount) {
                        return 1.0;
                    }
                }
            }

            double multiplier = inputTotal / outputTotal;
            return Math.round(multiplier * 100.0) / 100.0;

        } catch (Exception e) {
            e.printStackTrace();
            return 0;
        }
    }

    private static int extractPackCount(String text) {
        try {
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d+)\\s*pack").matcher(text);
            if (m.find()) {
                return Integer.parseInt(m.group(1));
            }
        } catch (Exception e) {
            // Ignore
        }
        return 0;
    }

    private static double convertToGrams(String qty, String unit) {
        double quantity = Double.parseDouble(qty);
        switch (unit.toLowerCase()) {
            case "kg":
                return quantity * 1000;
            case "g":
                return quantity;
            default:
                return quantity;
        }
    }

    private static double calculateTotalFromUom(String uom) {
        try {
            String[] parts = uom.split("[x*]");
            double total = 1.0;

            for (String part : parts) {
                part = part.trim();
                if (part.isEmpty()) {
                    continue;
                }

                java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d+\\.?\\d*)\\s*([a-zA-Z]*)").matcher(part);
                if (m.find()) {
                    String qty = m.group(1);
                    String unit = m.group(2);
                    if (unit.isEmpty()) {
                        unit = "g";
                    }
                    total *= convertToGrams(qty, unit);
                }
            }

            return total;
        } catch (Exception e) {
            e.printStackTrace();
            return 0;
        }
    }
}