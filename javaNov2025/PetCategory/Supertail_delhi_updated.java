package petCategory;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;



public class Supertail_delhi_updated {

    public static void main(String[] args) {
    	
    	ChromeOptions options = new ChromeOptions();
		options.addArguments("--headless"); // Run Chrome in headless mode
		options.addArguments("--disable-gpu"); // Disable GPU acceleration
		options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		options.addArguments("--start-maximized");	

		ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

		// Schedule the task to run every day at 7:00 AM
		Calendar now = Calendar.getInstance();
		Calendar nextRunTime = Calendar.getInstance();
		nextRunTime.set(Calendar.HOUR_OF_DAY, 3);
		nextRunTime.set(Calendar.MINUTE, 20);
		nextRunTime.set(Calendar.SECOND, 0);

		long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		if (initialDelay < 0) {
			initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		}

		scheduler.scheduleAtFixedRate(() -> {
			try {
				System.out.println("Starting web scraping task...");
				Supertail_delhi_updated.runWebScraping();
				System.out.println("Web scraping task completed.");
			} catch (Exception e) {
				e.printStackTrace();
			}
		}, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
	}
	public static void runWebScraping() throws Exception{

        
        WebDriver driver = new ChromeDriver();
        try {
            String filePath = ".\\input-data\\superTail Input data_Updated.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Data1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                    InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(), UOM = new ArrayList<>(),
                    Mulitiplier = new ArrayList<>(), Availability = new ArrayList<>(), Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();

            for (int i = 1; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                if (row == null) continue;

                String url = getCellValue(row, 5);
                if (!url.isEmpty()) {
                    inputPid.add(getCellValue(row, 0));
                    InputCity.add(getCellValue(row, 1));
                    InputName.add(getCellValue(row, 2));
                    InputSize.add(getCellValue(row, 3));
                    NewProductCode.add(getCellValue(row, 4));
                    uRL.add(url);
                    UOM.add(getCellValue(row, 6));
                    Mulitiplier.add(getCellValue(row, 7));
                    Availability.add(getCellValue(row, 8));
                    Pincode.add(getCellValue(row, 9));
                    NameForCheck.add(getCellValue(row, 10));
                }
            }

            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");
            createHeaderRow(resultsSheet);

            int rowIndex = 1;
            int ProductCount = 1;
            String currentPin = null;
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String uomStr = UOM.get(i);
                String availability = Availability.get(i);
                String locationSet = Pincode.get(i);
                String namecheck = NameForCheck.get(i);

                // Always recalculate multiplier
                String multiplierValue = calculateMultiplier(size, uomStr);
                System.out.println("Calculated Multiplier for " + size + "/" + uomStr + " = " + multiplierValue);

                try {
                    if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                        populateRow(resultsSheet, rowIndex++, id, city, name, size, "NA", "NA", "NA", "NA", "NA", multiplierValue, "NA", "NA", "NA", namecheck);
                        continue;
                    }

                    driver.get(url);
                    driver.manage().window().maximize();
                    if (currentPin == null || !currentPin.equals(locationSet)) {
                        selectLocation(driver, wait, locationSet);
                        currentPin = locationSet;
                    }

                    String pageTitle = driver.getTitle();
                    if (pageTitle.contains("404") || pageTitle.contains("Not Found")) {
                        populateRow(resultsSheet, rowIndex++, id, city, name, size, "NA", "NA", "NA", "NA", "NA", multiplierValue, "NA", "NA", "NA", namecheck);
                        continue;
                    }

                    String newName = getElementText(driver, By.xpath("//div[@class='grid__item medium-up--one-whole']//h1[@class='product-single__header h2']"), "NA");
                    System.out.println(newName);
                    String spRaw = getElementText(driver, By.xpath("//div[contains(@class, 'variant-input')]//input[@type='radio' and @checked]/following-sibling::label//div[contains(@class,'variant-box-middle')]/h1"), "NA");
                    String spValue = spRaw.equals("NA") ? "NA" : spRaw.replace("₹", "").replace(",", "");
                    System.out.println(spValue);
                    String mrpRaw = getElementText(driver, By.xpath("//div[contains(@class, 'variant-input')]//input[@type='radio' and @checked]/following-sibling::label//div[contains(@class,'variant-box-middle')]/p/span"), "NA");
                    String mrpValue = mrpRaw.equals("NA") ? "NA" : mrpRaw.replace("₹", "").replace(",", "");
                    System.out.println(mrpValue);
                    String uomValue1 = "NA";
                    try {
                        WebElement checkedInput = driver.findElement(By.xpath("//input[@type='radio' and @checked='checked']"));
                        WebElement label = checkedInput.findElement(By.xpath("./following-sibling::label"));
                        WebElement uomElement = label.findElement(By.xpath(".//div[contains(@class,'variant-box-header')]/p"));
                        uomValue1 = uomElement.getText().trim();
                        System.out.println(uomValue1);
                    } catch (Exception e) {
                        uomValue1 = "NA";
                        System.out.println(uomValue1);
                    }
                    
                    if (mrpValue.isEmpty()) mrpValue = spValue;
                    
                    int result = isAvailable(driver) ? 1 : 0;
                    String NewAvailability1 = String.valueOf(result);
                    System.out.println(NewAvailability1);
                    String offerValue = (mrpValue.equals(spValue)) ? "NA" : getOffer(driver);                    
                    System.out.println(offerValue);
                    takeScreenshot(driver, id);

                    System.out.println("---------------------------------" + ProductCount++ + " " + id + "------------------------");

                    populateRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url, newName, mrpValue, spValue, uomValue1, multiplierValue, NewAvailability1, offerValue, namecheck);

                } catch (Exception e) {
                    e.printStackTrace();
                    populateRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url, "NA", "NA", "NA", "NA", "NA", "NA", "NA", namecheck);
                }
            }

            String outputFilePath = ".\\Output\\SuperTail_OutputData_Delhi_" + new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss").format(new Date()) + ".xlsx";
            FileOutputStream outFile = new FileOutputStream(outputFilePath);
            resultsWorkbook.write(outFile);
            outFile.close();
            System.out.println("Output file saved: " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    public static String calculateMultiplier(String inputSize, String uom) {
        try {
            double inputSizeValue = parseSizeToGramsOrML(inputSize);
            double uomValue = parseSizeToGramsOrML(uom);
            if (uomValue == 0) return "NA";

            double multiplier = inputSizeValue / uomValue;
            double roundedUp = Math.ceil(multiplier * 10.0) / 10.0;
            return String.format("%.1f", roundedUp);
        } catch (Exception e) {
            return "NA";
        }
    }

    public static double parseSizeToGramsOrML(String sizeStr) {
        if (sizeStr == null || sizeStr.trim().isEmpty()) return 0;
        sizeStr = sizeStr.toLowerCase().replace("pack", "").replaceAll("[()\\s]", "");
        String[] parts = sizeStr.split("x");
        double total = 1.0;
        for (String part : parts) {
            total *= parseSingleUnit(part);
        }
        return total;
    }

    public static double parseSingleUnit(String unitStr) {
        unitStr = unitStr.toLowerCase().trim();
        if (unitStr.contains("kg")) {
            return Double.parseDouble(unitStr.replaceAll("[^\\d.]", "")) * 1000;
        } else if (unitStr.contains("ltr") || unitStr.contains("l")) {
            return Double.parseDouble(unitStr.replaceAll("[^\\d.]", "")) * 1000;
        } else if (unitStr.contains("g") || unitStr.contains("ml")) {
            return Double.parseDouble(unitStr.replaceAll("[^\\d.]", ""));
        } else {
            return Double.parseDouble(unitStr.replaceAll("[^\\d.]", ""));
        }
    }

    private static String getCellValue(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = { "InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP", "SP", "UOM", "Multiplier", "Availability", "Offer", "Commands", "Remarks", "Correctness", "Percentage", "Name", "Emp Id", "Name Check" };
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    private static void populateRow(Sheet sheet, int rowIndex, String id, String city, String name, String size, String productCode, String url, String newName, String mrp, String sp, String uom, String multiplier, String availability, String offer, String nameCheck) {
        Row row = sheet.createRow(rowIndex);
        String[] values = { id, city, name, size, productCode, url, newName, mrp, sp, uom, multiplier, availability, offer, "", "", "", "", "", "", nameCheck };
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    private static void selectLocation(WebDriver driver, WebDriverWait wait, String locationSet) {
        try {
            WebElement locationField = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='header-icons_wrapper']//div[@id='pin-code']")));
            locationField.click();
            WebElement popupInputBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@id='pincodeInput']")));
            popupInputBox.clear();
            popupInputBox.sendKeys(locationSet);
            WebElement applyButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='apply_btn']")));
            applyButton.click();
            Thread.sleep(2000);
        } catch (Exception e) {
            System.out.println("Error setting location: " + e.getMessage());
        }
    }

    private static boolean isAvailable(WebDriver driver) {
        try {
            WebElement addToCartBtn = driver.findElement(By.id("AddToCart-template--16703736905966__main"));
            return addToCartBtn.isEnabled();
        } catch (Exception e) {
            return false;
        }
    }

    private static String getOffer(WebDriver driver) {
        try {
            WebElement offer = driver.findElement(By.xpath("//div[contains(@class, 'variant-input')]//input[@type='radio' and @checked]/following-sibling::label//div[contains(@class,'variant-box-footer')]/p"));
            String text = offer.getText().replace("SAVE", "").trim();
            return text.endsWith("%") ? text + " Off" : text;
        } catch (Exception e) {
            return "NA";
        }
    }

    private static void takeScreenshot(WebDriver driver, String id) {
        try {
            TakesScreenshot ts = (TakesScreenshot) driver;
            File src = ts.getScreenshotAs(OutputType.FILE);
            String path = ".\\Screenshot\\SuperTail_" + id + "_" + new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss").format(new Date()) + ".png";
            FileUtils.copyFile(src, new File(path));
        } catch (Exception e) {
            System.out.println("Screenshot failed: " + e.getMessage());
        }
    }

    public static String getElementText(WebDriver driver, By locator, String defaultVal) {
        try {
            WebElement el = driver.findElement(locator);
            String text = el.getText().trim();
            return text.isEmpty() ? defaultVal : text;
        } catch (Exception e) {
            return defaultVal;
        }
    }
}
