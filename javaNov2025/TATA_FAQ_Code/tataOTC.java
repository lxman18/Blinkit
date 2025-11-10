package Code1;

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
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class tataOTC {
    public static void main(String[] args) {
        WebDriver driver = null;
        try {
            // Initialize ChromeDriver
            driver = new ChromeDriver();
            driver.manage().window().maximize();

            // Read URLs from Excel file
            String filePath = ".\\input-data\\sampleTata.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet3");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            List<String> inputName = new ArrayList<>(), urlList = new ArrayList<>();

            // Extract URLs and Names from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                if (i == 0 || row == null) {
                    continue;
                }

                Cell inputNameCell = row.getCell(0); // InputName
                Cell urlCell = row.getCell(1);       // URL

                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING)
                            ? inputNameCell.getStringCellValue()
                            : "";
                    inputName.add(name);
                    urlList.add(url);
                }
            }
            file.close();
            urlsWorkbook.close();

            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            // Create header row
            Row headerRow = resultsSheet.createRow(0);
            String[] headers = {
                    "InputName", "URL", "ProductName", "MRP", "SP", "STRIPE SIZE", "DELIVERY TIME", "Offer",
                    "Product highlights", "Key Ingredients", "Key Benefits", "Concerns It Can Help With", "General Description",
                    "Product Dimension", "Uses", "Product Specifications and Features", "Compatible With", "Indications",
                    "Directions for Use", "Safety Information", "Quick Tips", "Frequently Asked Questions",
                    "Other Information", "Expires on or after", "Product Form", "Good to Know", "Dosage",
                    "Effects of Deficiency", "Diet Type", "Suitable for", "Manufacturer details", "MARKETER NAME",
                    "Availability"
            };
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            int rowIndex = 1;
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

            for (int i = 0; i < urlList.size(); i++) {
                String name = inputName.get(i);
                String url = urlList.get(i);
                String newName = "NA";
                String mrpValue = "NA";
                String spValue = "NA";
                String offerValue = "NA";
                String stripeSize = "NA";
                String deliveryTime = "NA";
                String productHighlights = "NA";
                String keyIngredients = "NA";
                String keyBenefits = "NA";
                String concernsItCanHelpWith = "NA";
                String generalDescription = "NA";
                String productDimension = "NA";
                String uses = "NA";
                String productSpecificationsAndFeatures = "NA";
                String compatibleWith = "NA";
                String indications = "NA";
                String directionsForUse = "NA";
                String safetyInformation = "NA";
                String quickTips = "NA";
                String frequentlyAskedQuestions = "NA";
                String otherInformation = "NA";
                String expiresOnOrAfter = "NA";
                String productForm = "NA";
                String goodToKnow = "NA";
                String dosage = "NA";
                String effectsOfDeficiency = "NA";
                String dietType = "NA";
                String suitableFor = "NA";
                String manufacturerName = "NA";
                String marketerName = "NA";
                String availability = "In Stock";

                try {
                    driver.get(url);
                    if (i == 0) {
                        Thread.sleep(10000); // Initial wait for first page
                    }

                    // Product Name
                    try {
                        WebElement nameElement = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//h1 | //h1[@class='ProductTitle__product-title___3QMYH']")));
                        newName = nameElement.getText();
                        System.out.println("Product Name: " + newName);
                    } catch (Exception e) {
                        System.out.println("❌ Product name not found: " + e.getMessage());
                    }

                    // Selling Price (SP)
                    try {
                        WebElement sp = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//div[@class='PriceDetails__discount-div___nb724'] | //div[contains(@class, 'PriceDetails__selling-price')]/div")));
                        spValue = sp.getText().replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "").trim();
                        System.out.println("SP: " + spValue);
                    } catch (Exception e) {
                        System.out.println("⚠️ SP not found: " + e.getMessage());
                    }

                    // MRP
                    try {
                        WebElement mrp = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//span[@class='DiscountDetails__discount-price___Mdcwo'] | //div[contains(@class, 'PriceDetails__mrp')]/span[2]")));
                        mrpValue = mrp.getText().replace("₹", "").trim();
                        System.out.println("MRP: " + mrpValue);
                    } catch (Exception e) {
                        mrpValue = spValue; // Fallback to SP if MRP not found
                        System.out.println("⚠️ MRP not found: " + e.getMessage());
                    }

                    // Availability
                    try {
                        String xpathForTata = "//div[@class='OtcPriceBox__price-box___p13HY']//div[@class='OtcPriceBox__notify-me-wrapper___3Ckqb OtcPriceBox__fontSize14___5Uv2i']";
                        boolean isElementPresent = !driver.findElements(By.xpath(xpathForTata)).isEmpty();
                        availability = isElementPresent ? "Out of Stock" : "In Stock";
                        System.out.println("Availability: " + availability);
                    } catch (Exception e) {
                        System.out.println("⚠️ Availability check failed: " + e.getMessage());
                    }

                    // Offer
                    if (!mrpValue.equals(spValue) && !mrpValue.equals("NA") && !spValue.equals("NA")) {
                        try {
                            WebElement offer = driver.findElement(By.xpath("//span[@class='DiscountDetails__discount-percent___IfDdk']"));
                            offerValue = offer.getText().replace("% off", "% Off").trim();
                            System.out.println("Offer: " + offerValue);
                        } catch (Exception e) {
                            System.out.println("⚠️ Offer not found: " + e.getMessage());
                        }
                    }

                    // Stripe Size
                    try {
                        WebElement box = driver.findElement(By.xpath("//div[@class='DropdownA11y__display-text___QK-u8']"));
                        String quanBox = box.getText();
                        WebElement stripe = driver.findElement(By.xpath("//div[@class='OtcPriceBox__add-box___3rvCP']//span[3]"));
                        String quanStripe = stripe.getText();
                        stripeSize = quanBox + " of " + quanStripe;
                        System.out.println("Stripe Size: " + stripeSize);
                    } catch (Exception e) {
                        System.out.println("⚠️ Stripe Size not found: " + e.getMessage());
                    }

                    // Delivery Time
                    try {
                        WebElement delivery = driver.findElement(By.xpath("//div[@class='style__padded___2vNu9 style__headerText___3sw_C']//span[2]"));
                        deliveryTime = "Get in " + delivery.getText();
                        System.out.println("Delivery Time: " + deliveryTime);
                    } catch (Exception e) {
                        System.out.println("⚠️ Delivery Time not found: " + e.getMessage());
                    }

                    // Product Highlights
                    try {
                        List<WebElement> highlightLists = driver.findElements(By.xpath("//div[@class='ProductHighlights__highlights-text___dc-WQ']//ul/li"));
                        List<String> highlights = new ArrayList<>();
                        for (WebElement highlight : highlightLists) {
                            highlights.add(highlight.getText());
                        }
                        productHighlights = highlights.isEmpty() ? "NA" : String.join("\n\n", highlights);
                        System.out.println("Product Highlights: " + productHighlights);
                    } catch (Exception e) {
                        System.out.println("⚠️ Product Highlights not found: " + e.getMessage());
                    }

                    // Other Information
                    try {
                        List<WebElement> informationList = driver.findElements(By.xpath("//ul[@class='ProductSpecification__product-specification___GXpMu']//li"));
                        List<String> infoList = new ArrayList<>();
                        for (WebElement info : informationList) {
                            infoList.add(info.getText());
                        }
                        otherInformation = infoList.isEmpty() ? "NA" : String.join("\n\n", infoList);
                        System.out.println("Other Information: " + otherInformation);
                    } catch (Exception e) {
                        System.out.println("⚠️ Other Information not found: " + e.getMessage());
                    }

                    // Marketer Name
                    try {
                        WebElement marketer = driver.findElement(By.xpath(
                                "//h3[contains(text(), 'Marketer details')]/following-sibling::div[@class='OtcPage__compliance-info-wrapper___1edqX']//div[span[text()='Name: ']]"));
                        marketerName = marketer.getText().replace("Name: ", "").trim();
                        System.out.println("Marketer Name: " + marketerName);
                    } catch (Exception e) {
                        System.out.println("⚠️ Marketer Name not found: " + e.getMessage());
                    }

                    // Manufacturer Name
                    try {
                        JavascriptExecutor js = (JavascriptExecutor) driver;
                        js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
                        Thread.sleep(1000); // Wait for dynamic content
                        WebElement manufacturer = driver.findElement(By.xpath(
                                "//h3[contains(text(), 'Manufacturer details')]/following-sibling::div[@class='OtcPage__compliance-info-wrapper___1edqX']//div[span[text()='Name: ']]"));
                        manufacturerName = manufacturer.getText().replace("Name: ", "").trim();
                        System.out.println("Manufacturer Name: " + manufacturerName);
                    } catch (Exception e) {
                        System.out.println("⚠️ Manufacturer Name not found: " + e.getMessage());
                    }

                    // Description Parsing
                    try {
                        WebElement descriptionDiv = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//div[contains(@class, 'ProductDescription__product-description')]")));
                        String descriptionText = descriptionDiv.getText();
                        Map<String, String> parsed = parseDescriptionDynamically(descriptionText);

                        // Map parsed description to predefined headers
                        for (Map.Entry<String, String> entry : parsed.entrySet()) {
                            String key = entry.getKey();
                            String value = entry.getValue();
                            // Map to predefined headers (case-insensitive)
                            for (String header : headers) {
                                if (header.equalsIgnoreCase(key) || header.replace(" ", "").equalsIgnoreCase(key.replace(" ", ""))) {
                                    switch (header) {
                                        case "Key Ingredients":
                                            keyIngredients = value;
                                            break;
                                        case "Key Benefits":
                                            keyBenefits = value;
                                            break;
                                        case "Concerns It Can Help With":
                                            concernsItCanHelpWith = value;
                                            break;
                                        case "General Description":
                                            generalDescription = value;
                                            break;
                                        case "Product Dimension":
                                            productDimension = value;
                                            break;
                                        case "Uses":
                                            uses = value;
                                            break;
                                        case "Product Specifications and Features":
                                            productSpecificationsAndFeatures = value;
                                            break;
                                        case "Compatible With":
                                            compatibleWith = value;
                                            break;
                                        case "Indications":
                                            indications = value;
                                            break;
                                        case "Directions for Use":
                                            directionsForUse = value;
                                            break;
                                        case "Safety Information":
                                            safetyInformation = value;
                                            break;
                                        case "Quick Tips":
                                            quickTips = value;
                                            break;
                                        case "Frequently Asked Questions":
                                            frequentlyAskedQuestions = value;
                                            break;
                                        case "Expires on or after":
                                            expiresOnOrAfter = value;
                                            break;
                                        case "Product Form":
                                            productForm = value;
                                            break;
                                        case "Good to Know":
                                            goodToKnow = value;
                                            break;
                                        case "Dosage":
                                            dosage = value;
                                            break;
                                        case "Effects of Deficiency":
                                            effectsOfDeficiency = value;
                                            break;
                                        case "Diet Type":
                                            dietType = value;
                                            break;
                                        case "Suitable for":
                                            suitableFor = value;
                                            break;
                                    }
                                    System.out.println("Mapped " + key + " to " + header + ": " + value);
                                    break;
                                }
                            }
                        }
                        // Debug parsed description and raw description text
                        System.out.println("Parsed description for " + name + ": " + parsed);
                        if (!parsed.containsKey("Concerns It Can Help With") && !descriptionText.isEmpty()) {
                            System.out.println("Raw description for " + name + " (Concerns missing): \n\n" + descriptionText);
                        }
                    } catch (TimeoutException | NoSuchElementException e) {
                        System.out.println("❌ Description not found for: " + name);
                    } catch (Exception e) {
                        System.out.println("⚠️ Error parsing description for: " + name + " -> " + e.getMessage());
                    }

                    // Write to Excel
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(name);               // InputName
                    resultRow.createCell(1).setCellValue(url);                // URL
                    resultRow.createCell(2).setCellValue(newName);            // ProductName
                    resultRow.createCell(3).setCellValue(mrpValue);           // MRP
                    resultRow.createCell(4).setCellValue(spValue);            // SP
                    resultRow.createCell(5).setCellValue(stripeSize);         // STRIPE SIZE
                    resultRow.createCell(6).setCellValue(deliveryTime);       // DELIVERY TIME
                    resultRow.createCell(7).setCellValue(offerValue);         // Offer
                    resultRow.createCell(8).setCellValue(productHighlights);   // Product highlights
                    resultRow.createCell(9).setCellValue(keyIngredients);      // Key Ingredients
                    resultRow.createCell(10).setCellValue(keyBenefits);        // Key Benefits
                    resultRow.createCell(11).setCellValue(concernsItCanHelpWith); // Concerns It Can Help With
                    resultRow.createCell(12).setCellValue(generalDescription); // General Description
                    resultRow.createCell(13).setCellValue(productDimension);   // Product Dimension
                    resultRow.createCell(14).setCellValue(uses);               // Uses
                    resultRow.createCell(15).setCellValue(productSpecificationsAndFeatures); // Product Specifications and Features
                    resultRow.createCell(16).setCellValue(compatibleWith);     // Compatible With
                    resultRow.createCell(17).setCellValue(indications);        // Indications
                    resultRow.createCell(18).setCellValue(directionsForUse);   // Directions for Use
                    resultRow.createCell(19).setCellValue(safetyInformation);  // Safety Information
                    resultRow.createCell(20).setCellValue(quickTips);          // Quick Tips
                    resultRow.createCell(21).setCellValue(frequentlyAskedQuestions); // Frequently Asked Questions
                    resultRow.createCell(22).setCellValue(otherInformation);   // Other Information
                    resultRow.createCell(23).setCellValue(expiresOnOrAfter);   // Expires on or after
                    resultRow.createCell(24).setCellValue(productForm);        // Product Form
                    resultRow.createCell(25).setCellValue(goodToKnow);         // Good to Know
                    resultRow.createCell(26).setCellValue(dosage);             // Dosage
                    resultRow.createCell(27).setCellValue(effectsOfDeficiency); // Effects of Deficiency
                    resultRow.createCell(28).setCellValue(dietType);           // Diet Type
                    resultRow.createCell(29).setCellValue(suitableFor);        // Suitable for
                    resultRow.createCell(30).setCellValue(manufacturerName);   // Manufacturer details
                    resultRow.createCell(31).setCellValue(marketerName);       // MARKETER NAME
                    resultRow.createCell(32).setCellValue(availability);       // Availability

                    System.out.println("Data extracted for URL: " + url);

                } catch (Exception e) {
                    System.out.println("⚠️ Failed to extract data for URL: " + url + ": " + e.getMessage());
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(name);
                    resultRow.createCell(1).setCellValue(url);
                    resultRow.createCell(2).setCellValue(newName);
                    resultRow.createCell(3).setCellValue(mrpValue);
                    resultRow.createCell(4).setCellValue(spValue);
                    resultRow.createCell(5).setCellValue(stripeSize);
                    resultRow.createCell(6).setCellValue(deliveryTime);
                    resultRow.createCell(7).setCellValue(offerValue);
                    resultRow.createCell(8).setCellValue(productHighlights);
                    resultRow.createCell(9).setCellValue(keyIngredients);
                    resultRow.createCell(10).setCellValue(keyBenefits);
                    resultRow.createCell(11).setCellValue(concernsItCanHelpWith);
                    resultRow.createCell(12).setCellValue(generalDescription);
                    resultRow.createCell(13).setCellValue(productDimension);
                    resultRow.createCell(14).setCellValue(uses);
                    resultRow.createCell(15).setCellValue(productSpecificationsAndFeatures);
                    resultRow.createCell(16).setCellValue(compatibleWith);
                    resultRow.createCell(17).setCellValue(indications);
                    resultRow.createCell(18).setCellValue(directionsForUse);
                    resultRow.createCell(19).setCellValue(safetyInformation);
                    resultRow.createCell(20).setCellValue(quickTips);
                    resultRow.createCell(21).setCellValue(frequentlyAskedQuestions);
                    resultRow.createCell(22).setCellValue(otherInformation);
                    resultRow.createCell(23).setCellValue(expiresOnOrAfter);
                    resultRow.createCell(24).setCellValue(productForm);
                    resultRow.createCell(25).setCellValue(goodToKnow);
                    resultRow.createCell(26).setCellValue(dosage);
                    resultRow.createCell(27).setCellValue(effectsOfDeficiency);
                    resultRow.createCell(28).setCellValue(dietType);
                    resultRow.createCell(29).setCellValue(suitableFor);
                    resultRow.createCell(30).setCellValue(manufacturerName);
                    resultRow.createCell(31).setCellValue(marketerName);
                    resultRow.createCell(32).setCellValue(availability);
                }
            }

            // Save output Excel file
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = ".\\Output\\TAta1mg_New_output_" + timestamp + ".xlsx";
            FileOutputStream outFile = new FileOutputStream(outputFilePath);
            resultsWorkbook.write(outFile);
            outFile.close();
            resultsWorkbook.close();
            System.out.println("Output file saved: " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                System.out.println("DoNe DoNe Scraping DoNe");
                driver.quit();
            }
        }
    }

    public static Map<String, String> parseDescriptionDynamically(String text) {
        Map<String, String> map = new LinkedHashMap<>();
        String[] lines = text.split("\n\n");

        String currentHeader = "General Description";
        StringBuilder contentBuilder = new StringBuilder();
        boolean isCompatibleWithSection = false;
        boolean isProductSpecsSection = false;
        boolean isSafetyInformationSection = false;
        boolean isConcernsSection = false;

        // Define known headers to detect without relying solely on colon
        String[] knownHeaders = {
                "Key Ingredients", "Key Benefits", "Concerns It Can Help With", "Concerns",
                "General Description", "Product Dimension", "Uses", "Product Specifications and Features",
                "Compatible With", "Indications", "Directions for Use", "Safety Information",
                "Quick Tips", "Frequently Asked Questions", "Expires on or after", "Product Form",
                "Good to Know", "Dosage", "Effects of Deficiency", "Diet Type", "Suitable for"
        };

        for (String line : lines) {
            line = line.trim();
            if (line.isEmpty()) continue;

            String normalizedLine = line.replace(":", "").trim();
            boolean isHeader = false;
            String content = null;

            // Check for same-line headers (e.g., "Concerns It Can Help With: diaherria")
            for (String header : knownHeaders) {
                String headerLower = header.toLowerCase();
                if (normalizedLine.toLowerCase().startsWith(headerLower)) {
                    isHeader = true;
                    // Extract content after the header
                    String headerPart = header; // Use original header case
                    content = normalizedLine.substring(header.length()).trim();
                    if (content.startsWith(":") || content.startsWith("-")) {
                        content = content.substring(1).trim();
                    }
                    // Save previous section
                    if (contentBuilder.length() > 0) {
                        String contentSection = contentBuilder.toString().trim();
                        if (!contentSection.isEmpty()) {
                            String mappedHeader = currentHeader.equalsIgnoreCase("Concerns") ? "Concerns It Can Help With" : currentHeader;
                            map.put(mappedHeader, contentSection);
                        }
                    }
                    currentHeader = headerPart;
                    contentBuilder = new StringBuilder();
                    if (!content.isEmpty()) {
                        contentBuilder.append(content).append("\n\n");
                    }
                    break;
                }
            }

            // Fallback to multi-line header detection
            if (!isHeader) {
                isHeader = line.matches(".*:") && line.length() < 50;
                for (String header : knownHeaders) {
                    if (header.equalsIgnoreCase(normalizedLine)) {
                        isHeader = true;
                        break;
                    }
                }
                if (isHeader) {
                    // Save previous section
                    if (contentBuilder.length() > 0) {
                        String contentSection = contentBuilder.toString().trim();
                        if (!contentSection.isEmpty()) {
                            String mappedHeader = currentHeader.equalsIgnoreCase("Concerns") ? "Concerns It Can Help With" : currentHeader;
                            map.put(mappedHeader, contentSection);
                        }
                    }
                    currentHeader = normalizedLine;
                    contentBuilder = new StringBuilder();
                } else {
                    // Handle content based on section
                    if (isCompatibleWithSection && isProductSpecsSection) {
                        if (line.contains("Compatible With") || line.contains("Compatibility")) {
                            if (contentBuilder.length() > 0 && !map.containsKey("Product Specifications and Features")) {
                                map.put("Product Specifications and Features", contentBuilder.toString().trim());
                                contentBuilder = new StringBuilder();
                            }
                            contentBuilder.append(line).append("\n\n");
                            currentHeader = "Compatible With";
                            isProductSpecsSection = false;
                        } else {
                            contentBuilder.append(line).append("\n\n");
                            currentHeader = "Product Specifications and Features";
                            isCompatibleWithSection = false;
                        }
                    } else if (isSafetyInformationSection && line.contains("Safety Information")) {
                        contentBuilder.append(line).append("\n\n");
                    } else if (isConcernsSection) {
                        contentBuilder.append(line).append("\n\n");
                    } else {
                        contentBuilder.append(line).append("\n\n");
                    }
                }
            }

            // Track specific sections
            isCompatibleWithSection = currentHeader.equalsIgnoreCase("Compatible With");
            isProductSpecsSection = currentHeader.equalsIgnoreCase("Product Specifications and Features");
            isSafetyInformationSection = currentHeader.equalsIgnoreCase("Safety Information");
            isConcernsSection = currentHeader.equalsIgnoreCase("Concerns It Can Help With") || currentHeader.equalsIgnoreCase("Concerns");
        }

        // Save last section
        if (contentBuilder.length() > 0) {
            String content = contentBuilder.toString().trim();
            if (!content.isEmpty()) {
                String mappedHeader = currentHeader.equalsIgnoreCase("Concerns") ? "Concerns It Can Help With" : currentHeader;
                map.put(mappedHeader, content);
            }
        }

        return map;
    }
}

