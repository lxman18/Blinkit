import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.mail.*;
import javax.mail.internet.MimeMultipart;
import java.util.LinkedHashMap;

public class firstcry {
    // ---------- Fetch OTP from Gmail ----------
    public static String waitForOTP(String user, String password, int maxWaitSeconds) {
        String otp = null;
        int waited = 0;
        while (otp == null && waited < maxWaitSeconds) {
            otp = getLatestOTP(user, password);
            if (otp != null) break;
            try {
                Thread.sleep(5000);
                waited += 5;
                System.out.println("Waiting for OTP... " + waited + "s");
            } catch (InterruptedException e) {
                System.err.println("Interrupted while waiting for OTP: " + e.getMessage());
            }
        }
        return otp;
    }

    private static String getLatestOTP(String user, String password) {
        try {
            Properties props = new Properties();
            props.put("mail.store.protocol", "imaps");
            props.put("mail.imaps.host", "imap.gmail.com");
            props.put("mail.imaps.port", "993");
            props.put("mail.imaps.ssl.enable", "true");
            props.put("mail.imaps.timeout", "10000");

            Session session = Session.getInstance(props, null);
            Store store = session.getStore("imaps");
            store.connect("imap.gmail.com", user, password);

            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);

            Message[] messages = inbox.getMessages();
            if (messages.length == 0) {
                System.out.println("No emails found in inbox.");
                inbox.close(false);
                store.close();
                return null;
            }

            int start = Math.max(messages.length - 5, 0);
            for (int i = messages.length - 1; i >= start; i--) {
                Message message = messages[i];
                String from = message.getFrom()[0].toString();
                String subject = message.getSubject();

                if (from.contains("firstcry") || subject.toLowerCase().contains("otp")) {
                    System.out.println("Checking email from: " + from + ", Subject: " + subject);
                    String content = getTextFromMessage(message);

                    Pattern pattern = Pattern.compile("\\b\\d{6}\\b");
                    Matcher matcher = pattern.matcher(content);
                    if (matcher.find()) {
                        String otp = matcher.group(0);
                        System.out.println("Extracted OTP: " + otp);
                        inbox.close(false);
                        store.close();
                        return otp;
                    } else {
                        System.out.println("No OTP found in email from: " + from);
                    }
                }
            }

            System.out.println("No OTP email found in the latest " + (messages.length - start) + " emails.");
            inbox.close(false);
            store.close();
        } catch (AuthenticationFailedException e) {
            System.err.println("Gmail authentication failed. Verify email/password or App Password: " + e.getMessage());
        } catch (Exception e) {
            System.err.println("Error fetching OTP: " + e.getMessage());
            e.printStackTrace();
        }
        return null;
    }

    private static String getTextFromMessage(Message message) throws Exception {
        if (message.isMimeType("text/plain")) {
            return message.getContent().toString();
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            return getTextFromMimeMultipart(mimeMultipart);
        } else if (message.isMimeType("text/html")) {
            String html = message.getContent().toString();
            return html.replaceAll("<[^>]+>", "").replaceAll("\\s+", " ").trim();
        }
        return "";
    }

    private static String getTextFromMimeMultipart(MimeMultipart mimeMultipart) throws Exception {
        StringBuilder result = new StringBuilder();
        int count = mimeMultipart.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart bodyPart = mimeMultipart.getBodyPart(i);
            if (bodyPart.isMimeType("text/plain")) {
                result.append(bodyPart.getContent().toString());
            } else if (bodyPart.isMimeType("text/html")) {
                String html = bodyPart.getContent().toString();
                result.append(html.replaceAll("<[^>]+>", "").replaceAll("\\s+", " ").trim());
            } else if (bodyPart.getContent() instanceof MimeMultipart) {
                result.append(getTextFromMimeMultipart((MimeMultipart) bodyPart.getContent()));
            }
        }
        return result.toString();
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    public static void main(String[] args) throws Exception {
        final String GMAIL_USER = "firstcry1stbaby0369@gmail.com";
        final String GMAIL_APP_PASSWORD = "dwhtcunljwesoamf";

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--start-maximized");
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\Book1365.xlsx";
            List<String> inputPid = new ArrayList<>();
            List<String> InputCity = new ArrayList<>();
            List<String> InputName = new ArrayList<>();
            List<String> InputSize = new ArrayList<>();
            List<String> NewProductCode = new ArrayList<>();
            List<String> uRL = new ArrayList<>();
            List<String> UOM = new ArrayList<>();
            List<String> Multiplier = new ArrayList<>();
            List<String> Availability = new ArrayList<>();

            try (FileInputStream file = new FileInputStream(filePath);
                 Workbook urlsWorkbook = new XSSFWorkbook(file)) {
                Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
                if (urlsSheet == null) {
                    throw new IllegalArgumentException("Sheet 'FirstCry2' not found in the Excel file.");
                }

                int rowCount = urlsSheet.getPhysicalNumberOfRows();
                if (rowCount <= 1) {
                    System.out.println("No data rows found in the sheet.");
                    return;
                }

                for (int i = 1; i < rowCount; i++) {
                    Row row = urlsSheet.getRow(i);
                    if (row == null) {
                        System.out.println("Row " + i + " is empty, skipping.");
                        continue;
                    }

                    String id = getCellValueAsString(row.getCell(0));
                    String city = getCellValueAsString(row.getCell(1));
                    String name = getCellValueAsString(row.getCell(2));
                    String size = getCellValueAsString(row.getCell(3));
                    String productCode = getCellValueAsString(row.getCell(4));
                    String url = getCellValueAsString(row.getCell(5));
                    String uom = getCellValueAsString(row.getCell(6));
                    String multiplier = getCellValueAsString(row.getCell(7));
                    String availability = getCellValueAsString(row.getCell(8));

                    if (url != null && !url.trim().isEmpty() && !url.equalsIgnoreCase("NA")) {
                        inputPid.add(id);
                        InputCity.add(city);
                        InputName.add(name);
                        InputSize.add(size);
                        NewProductCode.add(productCode);
                        uRL.add(url);
                        UOM.add(uom);
                        Multiplier.add(multiplier);
                        Availability.add(availability);
                    } else {
                        System.out.println("Skipping row " + i + ": URL is empty or invalid.");
                    }
                }
            } catch (IOException e) {
                System.err.println("Error reading Excel file: " + e.getMessage());
                e.printStackTrace();
                return;
            }

            // Create Excel workbook for storing results
            try (Workbook resultsWorkbook = new XSSFWorkbook()) {
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

                // Perform login for the first URL
                if (!uRL.isEmpty()) {
                	driver.get("https://www.firstcry.com");	
        			WebDriverWait wait2 = new WebDriverWait(driver, Duration.ofSeconds(1));
        			WebElement pincode = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='geoLocation']")));
        			pincode.click();
        			WebElement serachinputpincode = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//sapn[@class='R14_link']")));
        			serachinputpincode.click();
        			WebElement pincodeinput = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='nonlpincode']")));
        			pincodeinput.click();
        			pincodeinput.clear();
        			pincodeinput.sendKeys("110015");
        			WebElement applybutton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//div[@class='appl-but']//span[@class='M16_white'])[1]")));
        			applybutton.click();
        			
        			//login 
        			WebElement loginbutton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@class='logreg']")));
        			loginbutton.click();
        			WebElement email = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[contains(@class,'cntryMobNo J14R_42')]")));
        			email.click();
        			email.sendKeys("firstcry1stbaby0369@gmail.com");
        			Thread.sleep(3000);
        			WebElement emaildrop = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='mailtip']")));
        			emaildrop.click();
        			Thread.sleep(2000);
        			
        			WebElement emailcontinue = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='loginotp']//span")));
        			emailcontinue.click();
        			
        			//Otp paste
//        			WebElement Otppaste = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div//input[@name='notp[]']")));
//        			Otppaste.click();
//        			
        			
                    System.out.println("Initiated OTP login.");
                    // Fetch and enter OTP
                    String otp = waitForOTP(GMAIL_USER, GMAIL_APP_PASSWORD, 90);
                    if (otp != null) {
                        System.out.println("Fetched OTP: " + otp);
                        List<WebElement> otpInputs = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div//input[@name='notp[]']")));
                        if (otpInputs.size() >= otp.length()) {
                            for (int i = 0; i < otp.length(); i++) {
                                WebElement otpField = otpInputs.get(i);
                                otpField.clear();
                                otpField.sendKeys(String.valueOf(otp.charAt(i)));
                            }
                            System.out.println("Entered OTP digits into fields.");
                        } else {
                            System.out.println("❌ Not enough OTP input fields found. Expected: " + otp.length() + ", Found: " + otpInputs.size());
                        }
                   Thread.sleep(2000);
                        WebElement submitOtp = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='SUBMIT' or text()='Login']")));
                        submitOtp.click();
                        System.out.println("Submitted OTP and logged in.");
                    } else {
                        System.out.println("❌ OTP not received within 90 seconds.");
                        return;
                    }
                }

                for (int i = 0; i < uRL.size(); i++) {
                    String id = inputPid.get(i);
                    String city = InputCity.get(i);
                    String name = InputName.get(i);
                    String size = InputSize.get(i);
                    String productCode = NewProductCode.get(i);
                    String url = uRL.get(i);
                    String uom = UOM.get(i);
                    String multiplier = Multiplier.get(i);
                    String availability = Availability.get(i);
                    String newName = "NA";
                    String mrpValue = "NA";
                    String finalSp = "NA";
                    String Offer1 = "NA";
                    String Offer2 = "NA";
                    String discount1 = "NA";
                    String discount2 = "NA";
                    String discount3 = "NA";
                    String sp1 = "NA";
                    String sp2 = "NA";
                    String sp3 = "NA";
                    String sp4 = "NA";
                    String sp5 = "NA";
                    String NewAvailability1 = availability;
                    String updateMultiplier = multiplier;

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

                        driver.get(url);
                        String oldFrame = "//section[@class='pinfosection']";
                        WebElement oldFrameCheck = null;

                        try {
                            oldFrameCheck = driver.findElement(By.xpath(oldFrame));
                        } catch (NoSuchElementException e1) {
                            System.out.println("Add to cart button not found.");
                        }

                        if (oldFrameCheck != null && oldFrameCheck.isDisplayed()) {
                            JavascriptExecutor js = (JavascriptExecutor) driver;
                            js.executeScript("window.scrollBy(0, 500)");

                            String addToCartButtonXPath1 = "/html/body/app-productdetail-rvp/span/section[1]/section/section[1]/div[4]/div/div[2]/span";
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
                                    System.out.println("Product Name: " + newName);
                                } catch (NoSuchElementException e) {
                                    try {
                                        WebElement nameElement = driver.findElement(By.xpath("//h1"));
                                        newName = nameElement.getText();
                                        System.out.println("Product Name: " + newName);
                                    } catch (Exception h) {
                                        try {
                                            WebElement nameElement = driver.findElement(By.xpath("//div[@class='right-contr']//div[@class='prod-info-wrap']//p[@class='prod-name R20_21']"));
                                            newName = nameElement.getText();
                                            System.out.println("Product Name: " + newName);
                                        } catch (Exception ex) {
                                            newName = "NA";
                                            System.out.println("Product Name not found.");
                                        }
                                    }
                                }

                                headercount++;
                                int Availability0 = 1;
                                NewAvailability1 = Integer.toString(Availability0);

                                try {
                                    WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                    mrpValue = mrp.getText();
                                    System.out.println("MRP: " + mrpValue);
                                } catch (NoSuchElementException e) {
                                    try {
                                        WebElement mrp = driver.findElement(By.xpath("//*[@id=\\\"prodImgInfo\\\"]/section[2]/section[1]/p/div/span/del"));
                                        mrpValue = mrp.getText();
                                        System.out.println("MRP: " + mrpValue);
                                    } catch (Exception o) {
                                        mrpValue = "NA";
                                        System.out.println("MRP not found.");
                                    }
                                }

                                Thread.sleep(2000);
                                List<WebElement> divElements = driver.findElements(By.xpath("//div[@class='dfpinner cpncode_block']//div[@class='dfpcoupan']"));
                                Map<String, List<String>> keyValueMap = new LinkedHashMap<>();
                                List<String> last7CharsList = new ArrayList<>();
                                int divCount = divElements.size();
                                System.out.println("Number of slides: " + divCount);

                                for (int j = 1; j <= divCount; j++) {
                                    try {
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
                                                    String last7Chars = "";
                                                    if (lastStarIndex != -1) {
                                                        String beforeLastStar = discountText.substring(0, lastStarIndex);
                                                        last7Chars = beforeLastStar.length() > 7 ? beforeLastStar.substring(beforeLastStar.length() - 7) : beforeLastStar;
                                                        System.out.println("Last 7 Characters Before Last '*': " + last7Chars);
                                                    }

                                                    String codeXpath = xpath + "//div[@class='cpninfo cpncode_btns']//div[@class='cpnname J13SB_42 cl_fff bg_29']";
                                                    WebElement codeElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(codeXpath)));
                                                    String code = codeElement.getText();
                                                    System.out.println("Coupon Code: " + code);

                                                    keyValueMap.computeIfAbsent(discountText, k -> new ArrayList<>()).add(code);
                                                    last7CharsList.add(last7Chars);
                                                }
                                            }
                                        }
                                    } catch (Exception e) {
                                        System.out.println("Error processing offer for slide " + j + ": " + e.getMessage());
                                    }
                                }

                                System.out.println("Final key-value pairs in the map:");
                                for (Map.Entry<String, List<String>> entry : keyValueMap.entrySet()) {
                                    String offerPercentage = entry.getKey();
                                    List<String> couponCodes = entry.getValue();
                                    System.out.println("Offer Percentage: " + offerPercentage + ", Coupon Codes: " + couponCodes);
                                }

                                try {
                                    driver.findElement(By.xpath("//div[@class='addgotoText btn_fill add_to_cart']//span")).click();
                                    Thread.sleep(1000);
                                    driver.findElement(By.xpath("//div[@class='addgotoText btn_fill go_to_cart']//span")).click();
                                } catch (NoSuchElementException e) {
                                    System.out.println("Error clicking add to cart or go to cart: " + e.getMessage());
                                }

                                Thread.sleep(2000);
                                String rateValue = "";
                                try {
                                    WebElement rate = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]")));
                                    String htmlContent = rate.getAttribute("innerHTML");
                                    rateValue = rate.getText().trim();

                                    if (htmlContent.contains("<sup>")) {
                                        double amount1 = Double.parseDouble(rateValue) / 100.0;
                                        finalSp = String.format("%.2f", amount1);
                                        System.out.println("Final SP (with sup): " + finalSp);
                                    } else {
                                        finalSp = rateValue;
                                        System.out.println("Final SP: " + finalSp);
                                    }
                                } catch (Exception e) {
                                    System.out.println("Error fetching rate: " + e.getMessage());
                                    finalSp = "NA";
                                }

                                // Commenting out BlinkitId.screenshot due to undefined class
                                /*
                                BlinkitId screenshot = new BlinkitId();
                                try {
                                    screenshot.screenshot(driver, "Firstcry", id);
                                } catch (Exception e) {
                                    System.out.println("Error taking screenshot: " + e.getMessage());
                                }
                                */

                                List<String> couponCodesList = new ArrayList<>();
                                List<String> offersList = new ArrayList<>(keyValueMap.keySet());
                                for (List<String> codes : keyValueMap.values()) {
                                    couponCodesList.addAll(codes);
                                }

                                Actions actions = new Actions(driver);
                                for (int p = 0; p < Math.min(offersList.size(), couponCodesList.size()); p++) {
                                    String couponCode = couponCodesList.get(p);
                                    String offerText = offersList.get(p);
                                    System.out.println("Applying Offer: " + offerText + ", Coupon Code: " + couponCode);

                                    try {
                                        WebElement coupon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='cupn_cod']//div[@class='input_field coup_inputfied div_input']//input")));
                                        actions.moveToElement(coupon).click().perform();
                                        coupon.clear();
                                        coupon.sendKeys(couponCode);

                                        WebElement applyClick = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[9]/div[2]/div[3]/div/span[2]")));
                                        actions.moveToElement(applyClick).click().perform();
                                        Thread.sleep(2000);

                                        WebElement rate1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]")));
                                        String htmlContent = rate1.getAttribute("innerHTML");
                                        rateValue = rate1.getText().trim();
                                        String formattedAmount;

                                        if (htmlContent.contains("<sup>")) {
                                            double amount = Double.parseDouble(rateValue) / 100.0;
                                            formattedAmount = String.format("%.2f", amount);
                                            System.out.println("Rate value for coupon code " + couponCode + ": " + formattedAmount);
                                        } else {
                                            formattedAmount = rateValue;
                                            System.out.println("Rate value for coupon code " + couponCode + ": " + formattedAmount);
                                        }

                                        String xpathExpression = "//div[@class='input_field coup_inputfied div_input']//p[@class='J12M_42 cl_e5 errmsg err1']";
                                        try {
                                            WebElement invalidCouponElement = driver.findElement(By.xpath(xpathExpression));
                                            System.out.println("Invalid coupon message displayed for code: " + couponCode);
                                        } catch (Exception g) {
                                            WebElement elements = driver.findElement(By.id("coponapply"));
                                            elements.click();
                                            System.out.println("'Invalid coupon' text not found for code: " + couponCode);
                                        }

                                        String extractOffer = offerText;
                                        if ("8% Off".equals(extractOffer) || "7% Off".equals(extractOffer)) {
                                            extractOffer = "5% Off";
                                        }

                                        switch (p) {
                                            case 0:
                                                Offer1 = extractOffer;
                                                sp1 = formattedAmount;
                                                break;
                                            case 1:
                                                Offer2 = extractOffer;
                                                sp2 = formattedAmount;
                                                break;
                                            case 2:
                                                discount1 = extractOffer;
                                                sp3 = formattedAmount;
                                                break;
                                            case 3:
                                                discount2 = extractOffer;
                                                sp4 = formattedAmount;
                                                break;
                                            case 4:
                                                discount3 = extractOffer;
                                                sp5 = formattedAmount;
                                                break;
                                            default:
                                                break;
                                        }

                                        System.out.println("Coupon after applying coupon " + (p + 1) + ": " + couponCode);
                                        System.out.println("SP after applying coupon " + (p + 1) + ": " + formattedAmount);
                                    } catch (Exception e) {
                                        System.out.println("Error applying coupon " + couponCode + ": " + e.getMessage());
                                    }
                                }

                                while (true) {
                                    try {
                                        WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='REMOVE']")));
                                        remove.click();
                                    } catch (Exception e) {
                                        try {
                                            WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']")));
                                            remove.click();
                                        } catch (Exception innerEx) {
                                            System.out.println("Error removing item: " + innerEx.getMessage());
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
                                int Availability0 = 0;
                                NewAvailability1 = Integer.toString(Availability0);

                                try {
                                    WebElement nameElement = driver.findElement(By.xpath("//h1[@class='J14M_42 cl_21 nonfastionpname']"));
                                    newName = nameElement.getText();
                                    System.out.println("Product Name: " + newName);
                                } catch (NoSuchElementException e) {
                                    try {
                                        WebElement nameElement = driver.findElement(By.xpath("//h1"));
                                        newName = nameElement.getText();
                                        System.out.println("Product Name: " + newName);
                                    } catch (Exception h) {
                                        newName = "NA";
                                        System.out.println("Product Name not found.");
                                    }
                                }

                                boolean isTextPresent = false;
                                try {
                                    String[] textsToCheck = {"NOTIFY ME"};
                                    String pageSource = driver.getPageSource();
                                    for (String text : textsToCheck) {
                                        if (pageSource.contains(text)) {
                                            isTextPresent = true;
                                            break;
                                        }
                                    }
                                } catch (Exception e) {
                                    System.out.println("Error checking page source: " + e.getMessage());
                                }

                                if (isTextPresent) {
                                    try {
                                        WebElement priceElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("prod-price")));
                                        String price = priceElement.getAttribute("data-price");
                                        finalSp = price;
                                        System.out.println("Notify me SP: " + finalSp);
                                    } catch (Exception e) {
                                        try {
                                            WebElement sp = driver.findElement(By.xpath("//span[@class='th-discounted-price ']//span"));
                                            finalSp = sp.getText();
                                            System.out.println("Notify me SP: " + finalSp);
                                        } catch (Exception ex) {
                                            finalSp = "NA";
                                            System.out.println("SP not found.");
                                        }
                                    }

                                    try {
                                        WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                        mrpValue = mrp.getText();
                                        System.out.println("MRP: " + mrpValue);
                                    } catch (NoSuchElementException e) {
                                        mrpValue = finalSp;
                                        System.out.println("MRP set to SP: " + mrpValue);
                                    }
                                } else {
                                    try {
                                        WebElement mrp = driver.findElement(By.xpath("//span[@class='J14R_42 cl_75']//del"));
                                        mrpValue = mrp.getText();
                                        System.out.println("MRP: " + mrpValue);
                                    } catch (NoSuchElementException e) {
                                        try {
                                            WebElement mrp = driver.findElement(By.xpath("//*[@id=\\\"prodImgInfo\\\"]/section[2]/section[1]/p/div/span/del"));
                                            mrpValue = mrp.getText();
                                            System.out.println("MRP: " + mrpValue);
                                        } catch (Exception ex) {
                                            mrpValue = "NA";
                                            System.out.println("MRP not found.");
                                        }
                                    }

                                    try {
                                        WebElement sp = driver.findElement(By.xpath("//span[@class='th-discounted-price ']//span"));
                                        String spValue = sp.getText();
                                        double amount1 = Double.parseDouble(spValue) / 100.0;
                                        finalSp = String.format("%.2f", amount1);
                                        System.out.println("Final SP: " + finalSp);
                                    } catch (Exception e) {
                                        finalSp = "NA";
                                        System.out.println("SP not found.");
                                    }
                                }
                            }
                        } else {
                            try {
                                WebElement nameElement = driver.findElement(By.id("prod_name"));
                                newName = nameElement.getText();
                                System.out.println("Product Name: " + newName);
                            } catch (NoSuchElementException e) {
                                try {
                                    WebElement nameElement = driver.findElement(By.xpath("//div[@class = 'prod-info-wrap']//following::p[1]"));
                                    newName = nameElement.getText();
                                    System.out.println("Product Name: " + newName);
                                } catch (Exception h) {
                                    try {
                                        WebElement nameElement = driver.findElement(By.xpath("//div[@class='right-contr']//div[@class='prod-info-wrap']//p[@class='prod-name R20_21']"));
                                        newName = nameElement.getText();
                                        System.out.println("Product Name: " + newName);
                                    } catch (Exception ex) {
                                        newName = "NA";
                                        System.out.println("Product Name not found.");
                                    }
                                }
                            }

                            headercount++;
                            int Availability0 = 1;
                            NewAvailability1 = Integer.toString(Availability0);

                            try {
                                WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                                mrpValue = mrp.getText();
                                System.out.println("MRP: " + mrpValue);
                            } catch (NoSuchElementException e) {
                                try {
                                    WebElement mrp = driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[2]/div[2]/div[2]/span[4]/span[3]"));
                                    mrpValue = mrp.getText();
                                    System.out.println("MRP: " + mrpValue);
                                } catch (Exception o) {
                                    try {
                                        WebElement mrp = driver.findElement(By.xpath("//span[@class='pos-rel2stat new-mrp-wrap']//span[@class='pmr R20_75 pos-rel2stat']"));
                                        mrpValue = mrp.getText();
                                        System.out.println("MRP: " + mrpValue);
                                    } catch (Exception ex) {
                                        mrpValue = "NA";
                                        System.out.println("MRP not found.");
                                    }
                                }
                            }

                            try {
                                driver.findElement(By.xpath("(//span[@class='step1 M16_white'])[1]//span")).click();
                                Thread.sleep(1000);
                                driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                            } catch (NoSuchElementException e) {
                                System.out.println("Error clicking add to cart or go to cart: " + e.getMessage());
                            }

                            Thread.sleep(2000);
                            String rateValue = "";
                            try {
                                WebElement rate = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[5]/div[4]/div[11]/div[4]/span[2]")));
                                rateValue = rate.getText();
                                double amount1 = Double.parseDouble(rateValue) / 100.0;
                                finalSp = String.format("%.2f", amount1);
                                System.out.println("Final SP: " + finalSp);
                            } catch (Exception e) {
                                System.out.println("Error fetching rate: " + e.getMessage());
                                finalSp = "NA";
                            }

                            while (true) {
                                try {
                                    WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='REMOVE']")));
                                    remove.click();
                                } catch (Exception e) {
                                    try {
                                        WebElement remove = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']")));
                                        remove.click();
                                    } catch (Exception innerEx) {
                                        System.out.println("Error removing item: " + innerEx.getMessage());
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
                        resultRow.createCell(10).setCellValue(updateMultiplier);
                        resultRow.createCell(11).setCellValue(NewAvailability1);
                        resultRow.createCell(12).setCellValue(Offer1);
                        resultRow.createCell(13).setCellValue(sp1);
                        resultRow.createCell(14).setCellValue(Offer2);
                        resultRow.createCell(15).setCellValue(sp2);
                        resultRow.createCell(16).setCellValue(discount1);
                        resultRow.createCell(17).setCellValue(sp3);
                        resultRow.createCell(18).setCellValue(discount2);
                        resultRow.createCell(19).setCellValue(sp4);
                        resultRow.createCell(20).setCellValue(discount3); // Fixed: was incorrectly discount2
                        resultRow.createCell(21).setCellValue(sp5);

                        System.out.println("Data extracted for URL: " + url);
                    } catch (Exception e) {
                        System.out.println("Failed to extract data for URL: " + url + " - " + e.getMessage());
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
                    }
                }

                // Save output file
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Firstcry_Diapers_OutputData_2_Half_" + timestamp + ".xlsx";
                try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
                    resultsWorkbook.write(outFile);
                    System.out.println("Output file saved: " + outputFilePath);
                } catch (IOException e) {
                    System.err.println("Error saving output file: " + e.getMessage());
                    e.printStackTrace();
                }
            }

        } catch (Exception e) {
            System.err.println("Error during execution: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
                System.out.println("Browser closed.");
            }
        }
    }
}