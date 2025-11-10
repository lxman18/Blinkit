package FirstCry;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import CommonUtility.BlinkitId;

public class Firstcry2halfbabayUpdated {
    public static void main(String[] args) throws Exception {
        WebDriver driver = new ChromeDriver();
        int count = 0;
        String spValue = "";
        String spValue1 = "";
        String finalSp = "";
        String newName = null;
        String mrpValue = null;
        String Offer2 = "NA";
        String Offer1 = "NA";
        String discount1 = "NA";
        String discount2 = "NA";
        String discount3 = "NA";
        String sp1 = "NA";
        String sp2 = "NA";
        String sp3 = "NA";
        String sp4 = "NA";
        String sp5 = "NA";
        String sp1Extra = "NA"; // New variable for SP 1 - Extra
        String sp2Extra = "NA"; // New variable for SP 2 - Extra
        String sp3Extra = "NA"; // New variable for SP 3 - Extra
        String sp4Extra = "NA"; // New variable for SP 4 - Extra
        String sp5Extra = "NA"; // New variable for SP 5 - Extra
        String NewAvailability1 = " ";
        String updateMulitipler = " ";
        String currentPin = null;
        String shippingFee = "NA";
        String platformFee = "NA";
        String netPayment = "NA";
        int extra = 54;

        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\firstcrynew.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("FirstCry2");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();
            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                    InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
                    UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), Availability = new ArrayList<>(),
                    Pincode = new ArrayList<>();

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
                Cell pincodeCell = row.getCell(9);
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
                    String pincode = (pincodeCell != null && pincodeCell.getCellType() == CellType.STRING)
                            ? pincodeCell.getStringCellValue()
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
                    Pincode.add(pincode);
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
            headerRow.createCell(10).setCellValue("MULTIPLIER");
            headerRow.createCell(11).setCellValue("AVAILABILITY");
            headerRow.createCell(12).setCellValue("OFFER 1");
            headerRow.createCell(13).setCellValue("SP 1");
            headerRow.createCell(14).setCellValue("OFFER 2");
            headerRow.createCell(15).setCellValue("SP 2");
            headerRow.createCell(16).setCellValue("DISCOUNT 1");
            headerRow.createCell(17).setCellValue("SP 3");
            headerRow.createCell(18).setCellValue("DISCOUNT 2");
            headerRow.createCell(19).setCellValue("SP 4");
            headerRow.createCell(20).setCellValue("DISCOUNT 3");
            headerRow.createCell(21).setCellValue("SP 5");
           

            int rowIndex = 1;
            int headercount = 0;

            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String uom = UOM.get(i);
                String mulitiplier = Mulitiplier.get(i);
                String availability = Availability.get(i);
                String pincode = Pincode.get(i);

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
                        resultRow.createCell(13).setCellValue("NA");
                        resultRow.createCell(14).setCellValue("NA");
                        resultRow.createCell(15).setCellValue("NA");
                        resultRow.createCell(16).setCellValue("NA");
                        resultRow.createCell(17).setCellValue("NA");
                        resultRow.createCell(18).setCellValue("NA");
                        resultRow.createCell(19).setCellValue("NA");
                        resultRow.createCell(20).setCellValue("NA");
                        resultRow.createCell(21).setCellValue("NA");
                        
                        System.out.println("Skipped processing for URL: " + url);
                        continue;
                    }

                    if (i == 0) {
                        driver.get("https://www.firstcry.com");
                        driver.manage().window().maximize();
                        Thread.sleep(5000);
                        WebElement reg = driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div[3]/ul/li[7]"));
                        reg.click();
                        Thread.sleep(5000);
                        WebElement regmail = driver.findElement(By.xpath("//*[@id=\"lemail\"]"));
                        regmail.click();
                        Thread.sleep(5000);
                        regmail.sendKeys("blktpoc2000@gmail.com");
                        Thread.sleep(3000);
                        WebElement conmail = driver.findElement(By.xpath("//*[@id='loginotp']//span"));
                        conmail.click();
                        Thread.sleep(30000);
                    }

                    driver.get(url);

                    if (currentPin == null || !currentPin.equals(pincode)) {
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//div[@class='pincodeCheck div_input changepintxt']")).click();
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//input[@id='lpincode']")).click();
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//input[@id='lpincode']")).clear();
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//input[@id='lpincode']")).sendKeys(pincode);
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("(//div[@class='appl-but']//span[.='APPLY'])[2]")).click();
                        Thread.sleep(2000);
                        currentPin = pincode;
                        driver.get(url);
                    }

                    String oldFrame = "//section[@class='pinfosection']";
                    WebElement oldFrameCheck = null;
                    try {
                        oldFrameCheck = driver.findElement(By.xpath(oldFrame));
                    } catch (NoSuchElementException e1) {
                        System.out.println("Add to cart button not found.");
                    }

                    if (oldFrameCheck != null && oldFrameCheck.isDisplayed()) {
                        JavascriptExecutor js = (JavascriptExecutor) driver;
                        js.executeScript("window.scrollBy(0, 300)");
                        Thread.sleep(2000);

                        String addToCartButtonXPath1 = "/html/body/app-productdetail-rvp/span/section[1]/section/section[1]/div[3]/div/div[2]/span";
                        WebElement addToCartButton = null;
                        try {
                            addToCartButton = driver.findElement(By.xpath(addToCartButtonXPath1));
                        } catch (NoSuchElementException e1) {
                            System.out.println("Add to cart button not found.");
                        }

                        if (addToCartButton != null && addToCartButton.isEnabled() && addToCartButton.isDisplayed()) {
                            System.out.println("Add to Cart button is present on the page.");
                            try {
                                WebElement nameElement = driver.findElement(By.xpath("//h1[@class='J14M_42 cl_21 nonfastionpname']"));
                                newName = nameElement.getText();
                                System.out.println(newName);
                            } catch (org.openqa.selenium.NoSuchElementException e) {
                                try {
                                    WebElement nameElement = driver.findElement(By.xpath("//h1"));
                                    newName = nameElement.getText();
                                    System.out.println(newName);
                                } catch (Exception h) {
                                    WebElement nameElement = driver.findElement(By.xpath("//div[@class='right-contr']//div[@class='prod-info-wrap']//p[@class='prod-name R20_21']"));
                                    newName = nameElement.getText();
                                    System.out.println(newName);
                                }
                            }

                            System.out.println("headercount = " + headercount);
                            headercount++;
                            int Availability0 = 1;
                            NewAvailability1 = Integer.toString(Availability0);

                            try {
                                WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                mrpValue = mrp.getText();
                                System.out.println(mrpValue);
                            } catch (org.openqa.selenium.NoSuchElementException e) {
                                try {
                                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"prodImgInfo\"]/section[2]/section[1]/p/div/span/del"));
                                    mrpValue = mrp.getText();
                                    System.out.println(mrpValue);
                                } catch (Exception o) {
                                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"prodImgInfo\"]/section[2]/section[1]/p/div/span/del"));
                                    mrpValue = mrp.getText();
                                    System.out.println(mrpValue);
                                }
                            }

                            Thread.sleep(2000);
                            List<WebElement> divElements = driver.findElements(By.xpath("//div[@class='dfpinner cpncode_block']//div[@class='dfpcoupan']"));
                            Map<String, List<String>> keyValueMap = new LinkedHashMap<>();
                            List<String> last7CharsList = new ArrayList<>();
                            int divCount = divElements.size();
                            System.out.println("Number of slides: " + divCount);
                            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                            String last7Chars = "";
                            String code = "";

                            for (int j = 1; j <= divCount; j++) {
                                try {
                                    Thread.sleep(3000);
                                    String xpath = "(//div[@class='dfpinner cpncode_block']//div[@class='dfpcoupan'])[" + j + "]";
                                    WebElement offerElement = wait.until(ExpectedConditions.visibilityOfElementLocated(
                                            By.xpath(xpath + "//div[@class='innersec']//span[@class='J14M_42 cpndsc']")));
                                    String offer = offerElement.getText();
                                    if (offer.contains("Extra") || offer.contains("FLAT") || offer.contains("extra") || offer.contains("Flat")) {
                                        String discountText = offer;
                                        if (discountText.contains("Non Club") || discountText.contains("FLAT") || discountText.contains("Extra") || discountText.contains("extra") || offer.contains("Flat")) {
                                            String[] parts = discountText.split(" ");
                                            int indexOfFlat = -1;
                                            int indexOfExtra = -1;
                                            for (int b = 0; b < parts.length; b++) {
                                                if (parts[b].equalsIgnoreCase("Flat")) {
                                                    indexOfFlat = b;
                                                }
                                                if (parts[b].equalsIgnoreCase("Extra")) {
                                                    indexOfExtra = b;
                                                }
                                            }
                                            if (indexOfFlat != -1 && indexOfFlat + 1 < parts.length) {
                                                String flatDiscountValue = parts[indexOfFlat + 1];
                                                System.out.println("Flat Discount Value: " + flatDiscountValue);
                                            }
                                            if (indexOfExtra != -1 && indexOfExtra + 1 < parts.length) {
                                                String extraDiscountValue = parts[indexOfExtra + 1];
                                                System.out.println("Extra Discount Value: " + extraDiscountValue);
                                            }
                                            if (discountText.contains("*")) {
                                                int lastStarIndex = discountText.lastIndexOf("*");
                                                if (lastStarIndex != -1) {
                                                    String beforeLastStar = discountText.substring(0, lastStarIndex);
                                                    last7Chars = beforeLastStar.length() > 7 ? beforeLastStar.substring(beforeLastStar.length() - 7) : beforeLastStar;
                                                    System.out.println("Last 7 Characters Before Last '*': " + last7Chars);
                                                }
                                                String codeXpath = xpath + "//div[@class='cpninfo cpncode_btns']//div[@class='cpnname J13SB_42 cl_fff bg_29']";
                                                WebElement codeElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(codeXpath)));
                                                code = codeElement.getText();
                                                System.out.println("Coupon Code: " + code);
                                                keyValueMap.computeIfAbsent(discountText, k -> new ArrayList<>()).add(code);
                                                last7CharsList.add(last7Chars);
                                            } else {
                                                System.out.println("No asterisk found in the text.");
                                            }
                                        } else {
                                            System.out.println("The text does not contain 'Non Club'.");
                                        }
                                    }
                                } catch (Exception e) {
                                    e.printStackTrace();
                                }
                            }

                            System.out.println("Final key-value pairs in the map:");
                            for (Map.Entry<String, List<String>> entry : keyValueMap.entrySet()) {
                                String offerPercentage = entry.getKey();
                                List<String> couponCodes = entry.getValue();
                                System.out.println("Offer Percentage: " + offerPercentage + ", Coupon Codes: " + couponCodes);
                            }

                            try {
                                Thread.sleep(1000);
                                driver.findElement(By.xpath("//div[@class='addgotoText btn_fill add_to_cart']//span")).click();
                                Thread.sleep(1000);
                                driver.findElement(By.xpath("//div[@class='addgotoText btn_fill go_to_cart']//span")).click();
                            } catch (NoSuchElementException e) {
                            }

                            Thread.sleep(2000);
                            String rateValue = "";
                            WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
                            Thread.sleep(4000);
                            wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]")));
                            WebElement rate = driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]"));
                            String htmlContent = rate.getAttribute("innerHTML");
                            rateValue = rate.getText().trim();
                            String formattedAmount;
                            if (htmlContent.contains("<sup>")) {
                                double amount1 = Double.parseDouble(rateValue) / 100.0;
                                formattedAmount = String.format("%.2f", amount1);
                                System.out.println("Final value for coupon code : " + formattedAmount);
                            } else {
                                formattedAmount = rateValue;
                            }

                            // New Logic: Extract Shipping Fee, Platform Fee, and Net Payment
                            try {
                                List<WebElement> shippingElements = driver.findElements(By.xpath("((//span[@class='J14R_42 cl_21 rht_val cl_7c'])[1]//span)[1]"));
                                if (!shippingElements.isEmpty()) {
                                    shippingFee = shippingElements.get(0).getText().replace("₹", "").trim();
                                    System.out.println("Shipping Fee: " + shippingFee);
                                } else {
                                    shippingFee = "0";
                                    System.out.println("Shipping Fee not found, defaulting to 0.");
                                }
                            } catch (Exception e) {
                                shippingFee = "0";
                                System.out.println("Shipping Fee not found, defaulting to 0.");
                            }

                            try {
                                List<WebElement> platformElements = driver.findElements(By.xpath("//span[@id='ConvenienceCharges']"));
                                if (!platformElements.isEmpty()) {
                                    platformFee = platformElements.get(0).getText().replace("₹", "").trim();
                                    System.out.println("Platform Fee: " + platformFee);
                                } else {
                                    platformFee = "0";
                                    System.out.println("Platform Fee not found, defaulting to 0.");
                                }
                            } catch (Exception e) {
                                platformFee = "0";
                                System.out.println("Platform Fee not found, defaulting to 0.");
                            }

                            try {
                                WebElement netPaymentElement = driver.findElement(By.xpath("//div[contains(text(),'Total Amount Payable')]/following-sibling::div/span"));
                                netPayment = netPaymentElement.getText().replace("₹", "").trim();
                                System.out.println("Net Payment: " + netPayment);
                            } catch (Exception e) {
                                System.out.println("Net Payment not found. Using fallback...");
                                netPayment = formattedAmount;
                                System.out.println("Fallback Net Payment: " + netPayment);
                            }

                            // Updated finalSp calculation
                            try {
                                double netPaymentValue = Double.parseDouble(netPayment);
                                double shippingFeeValue = shippingFee.equals("NA") ? 0.0 : Double.parseDouble(shippingFee);
                                double platformFeeValue = platformFee.equals("NA") ? 0.0 : Double.parseDouble(platformFee);
                                double finalSpValue = netPaymentValue - (shippingFeeValue + platformFeeValue);
                                finalSp = String.format("%.2f", finalSpValue);
                                System.out.println("Updated finalSp: " + finalSp);
                            } catch (NumberFormatException e) {
                                try {
                                    double netPaymentValue = Double.parseDouble(netPayment);
                                    double finalSpValue = netPaymentValue - extra;
                                    finalSp = String.format("%.2f", finalSpValue);
                                } catch (NumberFormatException ex) {
                                    finalSp = netPayment;
                                }
                                System.out.println("Error parsing values for finalSp calculation, using netPayment - extra: " + finalSp);
                            }

                            BlinkitId screenshot = new BlinkitId();
                            try {
                                BlinkitId.screenshot(driver, "Firstcry", id);
                            } catch (Exception e) {
                                e.fillInStackTrace();
                            }

                            updateMulitipler = mulitiplier;
                            Thread.sleep(3000);

                            List<String> couponCodesList = new ArrayList<>();
                            List<String> offersList = new ArrayList<>(keyValueMap.keySet());
                            for (List<String> codes : keyValueMap.values()) {
                                couponCodesList.addAll(codes);
                            }

                            Actions actions = new Actions(driver);
                            for (int p = 0; p < offersList.size(); p++) {
                                String couponCode = couponCodesList.get(p);
                                String offerText = last7CharsList.get(p);
                                System.out.println("Applying Offer: " + offerText + ", Coupon Code: " + couponCode);
                                WebElement coupon = driver.findElement(By.xpath("//div[@class='cupn_cod']//div[@class='input_field coup_inputfied div_input']//input "));
                                Thread.sleep(5000);
                                actions.moveToElement(coupon).click().perform();
                                Thread.sleep(500);
                                coupon.clear();
                                Thread.sleep(500);
                                coupon.sendKeys(couponCode);
                                Thread.sleep(1000);
                                wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]")));
                                WebElement applyClick = driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[9]/div[2]/div[3]/div/span[2]"));
                                actions.moveToElement(applyClick).click().perform();
                                Thread.sleep(6000);
                                WebElement rate1 = driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]"));
                                String formattedAmountCoupon;
                                htmlContent = rate1.getAttribute("innerHTML");
                                rateValue = rate1.getText().trim();
                                if (htmlContent.contains("<sup>")) {
                                    double amount = Double.parseDouble(rateValue) / 100.0;
                                    formattedAmountCoupon = String.format("%.2f", amount);
                                    System.out.println("Rate value for coupon code " + couponCode + ": " + formattedAmountCoupon);
                                } else {
                                    formattedAmountCoupon = rateValue;
                                }
                                Thread.sleep(2000);
                                String xpathExpression = "//div[@class='input_field coup_inputfied div_input']//p[@class='J12M_42 cl_e5 errmsg err1']";
                                try {
                                    WebElement invalidCouponElement = driver.findElement(By.xpath(xpathExpression));
                                    System.out.println("Invalid coupon message displayed.");
                                } catch (Exception g) {
                                    WebElement elements = driver.findElement(By.id("coponapply"));
                                    elements.click();
                                    System.out.println("'Invalid coupon' text not found on the webpage.");
                                }

                                // Calculate spX - extra
                                String spExtra = "NA";
                                if (!formattedAmountCoupon.equals("NA")) {
                                    try {
                                        double spValue11 = Double.parseDouble(formattedAmountCoupon);
                                        double spExtraValue = spValue11 - extra;
                                        spExtra = String.format("%.2f", spExtraValue);
                                        System.out.println("SP - Extra for coupon " + couponCode + ": " + spExtra);
                                    } catch (NumberFormatException e) {
                                        spExtra = "NA";
                                        System.out.println("Error calculating SP - Extra for coupon " + couponCode);
                                    }
                                }

                                String extractOffer = offerText;
                                if ("8% Off".equals(extractOffer) || "7% Off".equals(extractOffer)) {
                                    String offerText1 = "5% Off";
                                    offerText = offerText1;
                                }

                                if (offerText == null || couponCode == null || offerText.isEmpty() || couponCode.isEmpty()) {
                                    switch (p) {
                                        case 0:
                                            discount1 = "NA";
                                            sp3 = "NA";
                                            sp3Extra = "NA";
                                            break;
                                        case 1:
                                            discount2 = "NA";
                                            sp4 = "NA";
                                            sp4Extra = "NA";
                                            break;
                                        case 2:
                                            discount3 = "NA";
                                            sp5 = "NA";
                                            sp5Extra = "NA";
                                            break;
                                        case 3:
                                            Offer1 = "NA";
                                            sp1 = "NA";
                                            sp1Extra = "NA";
                                            break;
                                        case 4:
                                            Offer2 = "NA";
                                            sp2 = "NA";
                                            sp2Extra = "NA";
                                            break;
                                        default:
                                            break;
                                    }
                                } else {
                                    switch (p) {
                                        case 0:
                                            Offer1 = offerText;
                                            sp1 = formattedAmountCoupon;
                                            sp1Extra = spExtra;
                                            break;
                                        case 1:
                                            Offer2 = offerText;
                                            sp2 = formattedAmountCoupon;
                                            sp2Extra = spExtra;
                                            break;
                                        case 2:
                                            discount1 = offerText;
                                            sp3 = formattedAmountCoupon;
                                            sp3Extra = spExtra;
                                            break;
                                        case 3:
                                            discount2 = offerText;
                                            sp4 = formattedAmountCoupon;
                                            sp4Extra = spExtra;
                                            break;
                                        case 4:
                                            discount3 = offerText;
                                            sp5 = formattedAmountCoupon;
                                            sp5Extra = spExtra;
                                            break;
                                        default:
                                            break;
                                    }
                                }

                                System.out.println("Coupon after applying coupon " + (p + 1) + ": " + couponCode);
                                System.out.println("SP after applying coupon " + (p + 1) + ": " + formattedAmountCoupon);
                                System.out.println("SP - Extra after applying coupon " + (p + 1) + ": " + spExtra);
                            }

                            while (true) {
                                try {
                                    wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                                    WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='REMOVE']")));
                                    remove.click();
                                } catch (Exception e) {
                                    try {
                                        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                                        WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']")));
                                        remove.click();
                                    } catch (Exception innerEx) {
                                        innerEx.printStackTrace();
                                    }
                                }
                                try {
                                    Thread.sleep(2000);
                                    driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[3]/div/div[6]/div"));
                                } catch (NoSuchElementException e) {
                                    break;
                                }
                            }
                        } else {
                            System.out.println("Add to Cart button is NOT present on the page.");
                            Thread.sleep(2000);
                            int Availability1 = 1;
                            try {
                                WebElement nameElement = driver.findElement(By.xpath("//h1[@class='J14M_42 cl_21 nonfastionpname']"));
                                newName = nameElement.getText();
                                System.out.println(newName);
                            } catch (org.openqa.selenium.NoSuchElementException e) {
                                WebElement nameElement = driver.findElement(By.xpath("//h1"));
                                newName = nameElement.getText();
                                System.out.println(newName);
                            }

                            boolean isTextPresent = false;
                            try {
                                String[] textsToCheck = { "NOTIFY ME" };
                                String pageSource = driver.getPageSource();
                                for (String text : textsToCheck) {
                                    if (pageSource.contains(text)) {
                                        isTextPresent = true;
                                        break;
                                    }
                                }
                            } catch (Exception e) {
                                System.out.println(e.getMessage());
                            }

                            if (isTextPresent) {
                                try {
                                    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                                    WebElement priceElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("prod-price")));
                                    String price = priceElement.getAttribute("data-price");
                                    finalSp = price;
                                    System.out.println(finalSp);
                                } catch (Exception e) {
                                    WebElement sp = driver.findElement(By.xpath("//span[@class='th-discounted-price ']//span"));
                                    spValue = sp.getText();
                                    finalSp = spValue;
                                }
                                System.out.println("===Notify me sp ===" + finalSp);
                                try {
                                    WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                    mrpValue = mrp.getText();
                                    System.out.println(mrpValue);
                                } catch (org.openqa.selenium.NoSuchElementException e) {
                                    mrpValue = finalSp;
                                }
                                Thread.sleep(500);
                            } else {
                                try {
                                    WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                    mrpValue = mrp.getText();
                                    System.out.println(mrpValue);
                                } catch (org.openqa.selenium.NoSuchElementException e) {
                                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"prodImgInfo\"]/section[2]/section[1]/p/div/span/del"));
                                    mrpValue = mrp.getText();
                                    System.out.println(mrpValue);
                                }
                                Thread.sleep(500);
                                WebElement sp = driver.findElement(By.xpath("//span[@class='th-discounted-price ']//span"));
                                spValue = sp.getText();
                                double amount1 = Double.parseDouble(spValue) / 100.0;
                                String formattedAmount1 = String.format("%.2f", amount1);
                                System.out.println("Final value for coupon code : " + formattedAmount1);

                                // Updated finalSp calculation
                                try {
                                    double netPaymentValue = Double.parseDouble(formattedAmount1);
                                    double shippingFeeValue = shippingFee.equals("NA") ? 0.0 : Double.parseDouble(shippingFee);
                                    double platformFeeValue = platformFee.equals("NA") ? 0.0 : Double.parseDouble(platformFee);
                                    double finalSpValue = netPaymentValue - (shippingFeeValue + platformFeeValue);
                                    finalSp = String.format("%.2f", finalSpValue);
                                    System.out.println("Updated finalSp: " + finalSp);
                                } catch (NumberFormatException e) {
                                    try {
                                        double netPaymentValue = Double.parseDouble(formattedAmount1);
                                        double finalSpValue = netPaymentValue - extra;
                                        finalSp = String.format("%.2f", finalSpValue);
                                    } catch (NumberFormatException ex) {
                                        finalSp = formattedAmount1;
                                    }
                                    System.out.println("Error parsing values for finalSp calculation, using netPayment - extra: " + finalSp);
                                }
                            }
                            Availability1 = 0;
                            NewAvailability1 = Integer.toString(Availability1);
                            updateMulitipler = mulitiplier;
                            count = 1;
                        }
                    } else {
                        try {
                            WebElement nameElement = driver.findElement(By.id("prod_name"));
                            newName = nameElement.getText();
                            System.out.println(newName);
                        } catch (org.openqa.selenium.NoSuchElementException e) {
                            try {
                                WebElement nameElement = driver.findElement(By.xpath("//div[@class = 'prod-info-wrap']//following::p[1]"));
                                newName = nameElement.getText();
                                System.out.println(newName);
                            } catch (Exception h) {
                                WebElement nameElement = driver.findElement(By.xpath("//div[@class='right-contr']//div[@class='prod-info-wrap']//p[@class='prod-name R20_21']"));
                                newName = nameElement.getText();
                                System.out.println(newName);
                            }
                        }

                        System.out.println("headercount = " + headercount);
                        headercount++;
                        int Availability0 = 1;
                        NewAvailability1 = Integer.toString(Availability0);
                        updateMulitipler = mulitiplier;

                        try {
                            WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                            mrpValue = mrp.getText();
                            System.out.println(mrpValue);
                        } catch (org.openqa.selenium.NoSuchElementException e) {
                            try {
                                WebElement mrp = driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[2]/div[2]/div[2]/span[4]/span[3]"));
                                mrpValue = mrp.getText();
                                System.out.println(mrpValue);
                            } catch (Exception o) {
                                WebElement mrp = driver.findElement(By.xpath("//span[@class='pos-rel2stat new-mrp-wrap']//span[@class='pmr R20_75 pos-rel2stat']"));
                                mrpValue = mrp.getText();
                                System.out.println(mrpValue);
                            }
                        }

                        try {
                            Thread.sleep(1000);
                            driver.findElement(By.xpath("(//span[@class='step1 M16_white'])[1]//span")).click();
                            Thread.sleep(1000);
                            driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                        } catch (NoSuchElementException e) {
                        }

                        Thread.sleep(2000);
                        String rateValue = "";
                        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
                        Thread.sleep(4000);
                        wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[10]/div[4]/span[2]")));
                        WebElement rate = driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[10]/div[4]/span[2]"));
                        rateValue = rate.getText();
                        double amount1 = Double.parseDouble(rateValue) / 100.0;
                        String formattedAmount1 = String.format("%.2f", amount1);
                        System.out.println("Final value for coupon code : " + formattedAmount1);

                        // New Logic: Extract Shipping Fee, Platform Fee, and Net Payment
                        try {
                            List<WebElement> shippingElements = driver.findElements(By.xpath("//span[@id='ShippingCharges']"));
                            if (!shippingElements.isEmpty()) {
                                shippingFee = shippingElements.get(0).getText().replace("₹", "").trim();
                                System.out.println("Shipping Fee: " + shippingFee);
                            } else {
                                shippingFee = "0";
                                System.out.println("Shipping Fee not found, defaulting to 0.");
                            }
                        } catch (Exception e) {
                            shippingFee = "0";
                            System.out.println("Shipping Fee not found, defaulting to 0.");
                        }

                        try {
                            List<WebElement> platformElements = driver.findElements(By.xpath("//span[@id='ConvenienceCharges']"));
                            if (!platformElements.isEmpty()) {
                                platformFee = platformElements.get(0).getText().replace("₹", "").trim();
                                System.out.println("Platform Fee: " + platformFee);
                            } else {
                                platformFee = "0";
                                System.out.println("Platform Fee not found, defaulting to 0.");
                            }
                        } catch (Exception e) {
                            platformFee = "0";
                            System.out.println("Platform Fee not found, defaulting to 0.");
                        }

                        try {
                            WebElement netPaymentElement = driver.findElement(By.xpath("//div[contains(text(),'Total Amount Payable')]/following-sibling::div/span"));
                            netPayment = netPaymentElement.getText().replace("₹", "").trim();
                            System.out.println("Net Payment: " + netPayment);
                        } catch (Exception e) {
                            System.out.println("Net Payment not found. Using fallback...");
                            netPayment = formattedAmount1;
                            System.out.println("Fallback Net Payment: " + netPayment);
                        }

                        // Updated finalSp calculation
                        try {
                            double netPaymentValue = Double.parseDouble(netPayment);
                            double shippingFeeValue = shippingFee.equals("NA") ? 0.0 : Double.parseDouble(shippingFee);
                            double platformFeeValue = platformFee.equals("NA") ? 0.0 : Double.parseDouble(platformFee);
                            double finalSpValue = netPaymentValue - (shippingFeeValue + platformFeeValue);
                            finalSp = String.format("%.2f", finalSpValue);
                            System.out.println("Updated finalSp: " + finalSp);
                        } catch (NumberFormatException e) {
                            try {
                                double netPaymentValue = Double.parseDouble(netPayment);
                                double finalSpValue = netPaymentValue - extra;
                                finalSp = String.format("%.2f", finalSpValue);
                            } catch (NumberFormatException ex) {
                                finalSp = netPayment;
                            }
                            System.out.println("Error parsing values for finalSp calculation, using netPayment - extra: " + finalSp);
                        }

                        while (true) {
                            try {
                                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                                WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='REMOVE']")));
                                remove.click();
                            } catch (Exception e) {
                                try {
                                    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                                    WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']")));
                                    remove.click();
                                } catch (Exception innerEx) {
                                    innerEx.printStackTrace();
                                }
                            }
                            try {
                                Thread.sleep(2000);
                                driver.findElement(By.xpath("/html/body/form/section[1]/section[5]/div[3]/div/div[6]/div"));
                            } catch (NoSuchElementException e) {
                                break;
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
                    resultRow.createCell(8).setCellValue(finalSp);
                    resultRow.createCell(9).setCellValue(uom);
                    resultRow.createCell(10).setCellValue(updateMulitipler);
                    resultRow.createCell(11).setCellValue(NewAvailability1);
                    resultRow.createCell(12).setCellValue(Offer1);
                    resultRow.createCell(13).setCellValue(sp1Extra);
                    resultRow.createCell(14).setCellValue(Offer2);
                    resultRow.createCell(15).setCellValue(sp2Extra);
                    resultRow.createCell(16).setCellValue(discount1);
                    resultRow.createCell(17).setCellValue(sp3Extra);
                    resultRow.createCell(18).setCellValue(discount2);
                    resultRow.createCell(19).setCellValue(sp4Extra);
                    resultRow.createCell(20).setCellValue(discount3);
                    resultRow.createCell(21).setCellValue(sp5Extra);
                    
                    System.out.println("Data extracted for URL: " + url);

                    // Reset variables for the next iteration
                    Offer1 = "NA";
                    sp1 = "NA";
                    sp1Extra = "NA";
                    Offer2 = "NA";
                    sp2 = "NA";
                    sp2Extra = "NA";
                    discount1 = "NA";
                    sp3 = "NA";
                    sp3Extra = "NA";
                    discount2 = "NA";
                    sp4 = "NA";
                    sp4Extra = "NA";
                    discount3 = "NA";
                    sp5 = "NA";
                    sp5Extra = "NA";
                    shippingFee = "NA";
                    platformFee = "NA";
                    netPayment = "NA";
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
                    resultRow.createCell(9).setCellValue("NA");
                    resultRow.createCell(10).setCellValue("NA");
                    resultRow.createCell(11).setCellValue("NA");
                    resultRow.createCell(12).setCellValue("NA");
                    resultRow.createCell(13).setCellValue("NA");
                    resultRow.createCell(14).setCellValue("NA");
                    resultRow.createCell(15).setCellValue("NA");
                    resultRow.createCell(16).setCellValue("NA");
                    resultRow.createCell(17).setCellValue("NA");
                    resultRow.createCell(18).setCellValue("NA");
                    resultRow.createCell(19).setCellValue("NA");
                    resultRow.createCell(20).setCellValue("NA");
                    resultRow.createCell(21).setCellValue("NA");
                    resultRow.createCell(22).setCellValue("NA");
                    resultRow.createCell(23).setCellValue("NA");
                    resultRow.createCell(24).setCellValue("NA");
                    resultRow.createCell(25).setCellValue("NA");
                    resultRow.createCell(26).setCellValue("NA");
                    resultRow.createCell(27).setCellValue("NA");
                    resultRow.createCell(28).setCellValue("NA");
                    resultRow.createCell(29).setCellValue("NA");
                    System.out.println("Failed to extract data for URL: " + url);
                }
            }

            try {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Firstcry_Baby_2_HALF_OutputData_" + timestamp + ".xlsx";
                FileOutputStream outFile = new FileOutputStream(outputFilePath);
                resultsWorkbook.write(outFile);
                outFile.close();
                System.out.println("Output file saved: " + outputFilePath);
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }
}