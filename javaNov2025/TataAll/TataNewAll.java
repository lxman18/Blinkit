package Code1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeoutException;

import javax.print.DocFlavor.STRING;

public class TataNewAll {
    public static void main(String[] args) throws Exception {
        // Initialize ChromeOptions
    	
        ChromeOptions options = new ChromeOptions();
        
        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        List<String> InputName = new ArrayList<>(), uRL = new ArrayList<>();

        Workbook resultsWorkbook = new XSSFWorkbook();
        Sheet resultsSheet = resultsWorkbook.createSheet("Results");
        createHeaderRow(resultsSheet);

        int rowIndex = 1;
        String currentPin = null;

        try (FileInputStream file = new FileInputStream(".\\input-data\\Book3.xlsx");
             Workbook urlsWorkbook = new XSSFWorkbook(file)) {

            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            // Extract URLs from Excel
            for (int i = 1; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                if (row.getCell(1) != null) {
                    String url;
                    if (row.getCell(1).getCellType() == CellType.STRING) {
                        url = row.getCell(1).getStringCellValue();
                    } else if (row.getCell(1).getCellType() == CellType.NUMERIC) {
                        url = String.valueOf(row.getCell(1).getNumericCellValue());
                    } else {
                        url = "NA";
                    }
                    InputName.add(row.getCell(0).getStringCellValue());
                    uRL.add(url);
                }
            }

            int ProductCOUNT = 0;
            int headercount = 0; // Declare headercount

            // Main data extraction logic
            for (int i = 0; i < uRL.size(); i++) {
                String url = uRL.get(i);
                if (url == null || url.isEmpty() || url.equalsIgnoreCase("NA")) {
                    writeResults(resultsSheet, rowIndex++, InputName.get(i), url, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA","NA","NA","NA","NA","NA");
                    System.out.println("Skipped processing for URL: " + url);
                    continue;
                }

                try {
                    driver.get(url);
                    Thread.sleep(2000);
                    
                    scrollToBottom(driver);

                    System.out.println("PRODUCTCOUNT = " + ProductCOUNT);
                    ProductCOUNT++;

                    String Frame = "//div[@class='DrugPage__wrapper___3olBc']";
                    WebElement FrameCheck = null;

                    try {
                    	FrameCheck = driver.findElement(By.xpath(Frame));
                    } catch (NoSuchElementException e1) {
                        System.out.println("Old Code is Executing...");
                    }

                    // Declare all variables at the top to avoid scope issues
                    String newName = "NA", mrpValue = "NA", spValue = "NA", offerValue = "NA", stripSize = "NA",
                            delTime = "NA", marketName = "NA", saltName = "NA", storageName = "NA", proInfo = "NA",
                            useInfo = "NA", benefits = "NA", sideEffects = "NA", prescription = "NA", HowtoUse="NA", Howitwork="NA",Safety="NA",Interaction="NA",Quesstion="NA";
                    String originalSp1 = "NA", originalMrp1 = "NA";

                    if (FrameCheck != null && FrameCheck.isDisplayed()) {
                        try {
                            newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))).getText();
                            System.out.println("newName =" + newName);
                        } catch (Exception e) {
                            newName = "NA";
                        }

                        Thread.sleep(3000);

                        mrpValue = extractMRP(driver);
                        System.out.println("mrpValue = " + mrpValue);
                        spValue = extractSP(driver);
                        System.out.println("spValue = " + spValue);
                        offerValue = extractOffer(driver);
                        System.out.println("offerValue = " + offerValue);
                        stripSize = extractStripSize(driver);
                        System.out.println("stripSize = " + stripSize);
                        delTime = extractDelTime(driver);
                        System.out.println("delTime = " + delTime);
                        marketName = extractMarketer(driver);
                        System.out.println("marketName = " + marketName);
                        saltName = extractSalt(driver);
                        System.out.println("saltName =" + saltName);
                        storageName = extractStorage(driver);
                        System.out.println("storageName = " + storageName);
                        proInfo = extractProductInfo(driver);
                        System.out.println("proInfo = " + proInfo);
                        useInfo = extractUsesInfo(driver);
                        System.out.println("useInfo = " + useInfo);
                        benefits = extractBenefits(driver);
                        System.out.println("benefits = " + benefits);
                        sideEffects = extractSideEffects(driver);
                        System.out.println("sideEffects = " + sideEffects);
                        prescription = extractprescription(driver);
                        System.out.println("Prescription = " + prescription);
                        HowtoUse=HowUse(driver);
                        System.out.println("How To Use =" + HowtoUse);
                        Howitwork=HowWork(driver);
                        System.out.println("How it Works =" + Howitwork);
                        Safety=safetyAdvice(driver);
                        System.out.println("Safety Advice =" +Safety);
                        Interaction=scrapeDrugInteractions(driver);
                        System.out.println("Interaction with Drugs =" + Interaction);
                        Quesstion=scrapeFaq(driver);
                        System.out.println("FAQs =" + Quesstion);

                    } else {
                    	try {

    						WebElement nameElement = driver.findElement(By.xpath("//h1"));
    						newName = nameElement.getText();
    						System.out.println(newName);
                    		}

    					catch(NoSuchElementException e) {

    						WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/div/div/div/div/div[2]/div[3]/div[1]/div[1]/h1"));
    						newName = nameElement.getText();
    						System.out.println(newName);

    							}
                    	
                      
                        Thread.sleep(5000);

                        try {
    						WebElement sp = driver.findElement(By.xpath("//div[@class='PriceDetails__discount-div-sale___1aQb9 PriceDetails__sale-price___1OfJd']"));
    						originalSp1 = sp.getText();
    						spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "");
    						System.out.println(spValue);

    					}
    					catch(Exception e) {
    						try {
    							WebElement sp = driver.findElement(By.xpath("//*[@id='container']/div/div/div[2]/div[4]/div[1]/div/div[2]/div[2]/div"));
    							originalSp1 = sp.getText();
    							spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "");
    							System.out.println(spValue);
    						}
    						catch (Exception e2) {
    							spValue = "NA";
    						}}
                        Thread.sleep(2000);
                        
                    	Thread.sleep(2000);
    					try {
    						WebElement mrp = driver.findElement(By.xpath("//span[@class='DiscountDetails__discount-price___Mdcwo']"));
    						originalMrp1 = mrp.getText();
    						mrpValue = originalMrp1.replace("₹", "");
    						System.out.println(mrpValue);

    					} 

    					catch(NoSuchElementException e){
    						mrpValue = spValue;
    					}}

                    // Write results to the results sheet
                    writeResults(resultsSheet, rowIndex++, InputName.get(i), url, newName, mrpValue, spValue, stripSize,
                            delTime, marketName, saltName, storageName, proInfo, useInfo, benefits, sideEffects, offerValue, prescription,HowtoUse,Howitwork,Safety,Interaction,Quesstion);

                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Failed to extract data for URL: " + url);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("An error occurred during the extraction process.");
        } finally {
            saveResultsToExcel(resultsWorkbook);
            if (driver != null) {
                System.out.println("Closing the driver.");
                driver.quit();
            }
        }
    }

    private static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"INPUT NAME", "URL", "PRODUCT NAME", "MRP", "SP", "STRIPE SIZE", "DELIVERY TIME", "MARKETER NAME",
                "SALT", "STORAGE", "PRODUCT INFO", "USES", "BENEFITS", "SIDE EFFECTS", "OFFER", "PRESCRIPTION","HOW TO USE","HOW IT WORKS","SAFETY ADVICE","INTERACTION WITH DRUGS","FAQs"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    public static void writeResults(Sheet resultsSheet, int rowIndex, String inputName, String url, String newName,
                                    String mrpValue, String spValue, String stripSize, String delTime, String marketName,
                                    String saltName, String storageName, String proInfo, String useInfo, String benefits,
                                    String sideEffects, String offerValue, String prescription, String HowtoUse, String Howitwork, String Safety, String Interaction,String Quesstion  ) {
        Row row = resultsSheet.createRow(rowIndex);
        row.createCell(0).setCellValue(inputName);
        row.createCell(1).setCellValue(url);
        row.createCell(2).setCellValue(newName);
        row.createCell(3).setCellValue(mrpValue);
        row.createCell(4).setCellValue(spValue);
        row.createCell(5).setCellValue(stripSize);
        row.createCell(6).setCellValue(delTime);
        row.createCell(7).setCellValue(marketName);
        row.createCell(8).setCellValue(saltName);
        row.createCell(9).setCellValue(storageName);
        row.createCell(10).setCellValue(proInfo);
        row.createCell(11).setCellValue(useInfo);
        row.createCell(12).setCellValue(benefits);
        row.createCell(13).setCellValue(sideEffects);
        row.createCell(14).setCellValue(offerValue);
        row.createCell(15).setCellValue(prescription);
        row.createCell(16).setCellValue(HowtoUse);
        row.createCell(17).setCellValue(Howitwork);
        row.createCell(18).setCellValue(Safety);
        row.createCell(19).setCellValue(Interaction);
        row.createCell(20).setCellValue(Quesstion);
        
    }

    private static void saveResultsToExcel(Workbook resultsWorkbook) {
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = ".\\Output\\TATA_OutputData_sheet1_" + timestamp + ".xlsx";
            try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
                resultsWorkbook.write(outFile);
            }
            System.out.println("Output file saved: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("Failed to save the output file.");
        }
    }

    private static String extractMRP(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//span[@class='PriceBoxPlanOption__margin-right-4___2aqFt PriceBoxPlanOption__stike___pDQVN']"));
            if (elements.isEmpty()) {
                try {
                    elements = driver.findElements(By.xpath("//div[.='MRP']/following-sibling::div[1]"));
                } catch (Exception e) {
                    elements = driver.findElements(By.xpath("//span[.='MRP']/following-sibling::span[1]"));
                }
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractSP(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//span[@class='PriceBoxPlanOption__offer-price___3v9x8 PriceBoxPlanOption__offer-price-cp___2QPU_'])[1]"));
            if (elements.isEmpty()) {
                try {
                    elements = driver.findElements(By.xpath("//div[.='MRP']/following-sibling::div[1]"));
                } catch (Exception e) {
                    elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__price___dj2lv']//div[2]"));
                }
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractOffer(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//span[@class='PriceBoxPlanOption__discount___iN_jm']"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//span[.='MRP']/following-sibling::span[2]"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractStripSize(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugPriceBox__qty-wrapper___1RBzv']//div[2])[2]"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractDelTime(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='style__box___1ez55']//div)[1]"));
            if (elements.isEmpty()) {
                String pageSource = driver.getPageSource();
                String[] textsToCheck = {"SOLD OUT", "DISCONTINUED"};
                for (String text : textsToCheck) {
                    if (pageSource.contains(text)) {
                        return "SOLD OUT or DISCONTINUED";
                    }
                }
                return "NA";
            } else {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractMarketer(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//div[.='Marketer']/following-sibling::div"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractSalt(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//div[.='SALT COMPOSITION']/following-sibling::div"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractStorage(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//div[.='Storage']/following-sibling::div"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractProductInfo(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX'])[1]"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractUsesInfo(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX']//ul)[1]"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractBenefits(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX']//div[@class='ShowMoreArray__tile___2mFZk'])[1]"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractSideEffects(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//div[@id='side_effects']//div[@class='DrugOverview__list-container___2eAr6 DrugOverview__content___22ZBX']/ul"));
            if (elements.isEmpty()) {
                elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            }
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static String extractprescription(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//span[text()='Prescription Required']"));
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }

    private static boolean isValidValue(String value) {
        return value != null && !value.isEmpty() && !value.equals("₹");
    }
    
    private static String HowUse(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX'])[6]"));
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }
    
    private static String HowWork(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX'])[7]"));
            if (!elements.isEmpty()) {
                String text = elements.get(0).getText();
                return isValidValue(text) ? text.trim() : "NA";
            } else {
                return "NA";
            }
        } catch (Exception e) {
            return "NA";
        }
    }
    
    private static String safetyAdvice(WebDriver driver) {
        try {
            // Get all the warning blocks
            List<WebElement> titles = driver.findElements(By.xpath("//div[@id='safety_advice']//div[contains(@class,'DrugOverview__warning-top')]"));
            StringBuilder safetyAdvice = new StringBuilder();

            for (int i = 0; i < titles.size(); i++) {
                WebElement titleElement = titles.get(i);
                try {
                    // Extract the title (e.g., Alcohol, Pregnancy)
                    String title = titleElement.findElement(By.xpath("./span")).getText();

                    // Extract the warning level (e.g., SAFE, UNSAFE)
                    String levelText = "";
                    try {
                        WebElement levelElement = titleElement.findElement(By.xpath(".//div[contains(@class,'DrugOverview__warning-tag')]"));
                        if (levelElement.isDisplayed()) {
                            levelText = levelElement.getText();
                            System.out.println("Level: " + levelText);
                        }
                    } catch (Exception e) {
                        System.out.println("Level not found: " + e.getMessage());
                    }


                    // Find the corresponding advice block
                    WebElement adviceElement = titleElement.findElement(By.xpath("following-sibling::div[contains(@class,'DrugOverview__content')]"));
                    String advice = adviceElement.getText();

                    safetyAdvice.append("Title: ").append(title).append("\n")
                                .append("Level: ").append(levelText).append("\n")
                                .append("Advice: ").append(advice).append("\n")
                                .append("=====================================\n");

                } catch (NoSuchElementException e) {
                    safetyAdvice.append("Element not found for warning at index ").append(i).append("\n");
                }
            }

            return safetyAdvice.toString();

        } catch (Exception e) {
            return "Error occurred: " + e.getMessage();
        }
    }

    public static String scrapeDrugInteractions(WebDriver driver) throws Exception {
        String interactionsData = "";
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        String descriptionText="";
        try {
            // Wait for interaction elements to be present
            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='DrugInteraction__drug___1XyzI']")));
            List<WebElement> interactionElements = driver.findElements(By.xpath("//div[@class='DrugInteraction__drug___1XyzI']"));
            System.out.println("Found " + interactionElements.size() + " interaction elements.");

            if (interactionElements.isEmpty()) {
                return "No drug interaction elements found.\n";
            }

            for (int i = 0; i < interactionElements.size(); i++) {
                WebElement interactionElement = interactionElements.get(i);

                try {
                    // Extract drug name
                    String drugName = interactionElement.findElement(By.xpath(".//span[@class='DrugInteraction__drug-interaction-name___2jI1z']")).getText();

                    // Extract interaction level
                    String interactionLevel = "";
                    try {
                        interactionLevel = interactionElement.findElement(By.xpath(".//span[contains(@class, 'DrugInteraction__moderate')]")).getText();
                    } catch (NoSuchElementException e) {
                        interactionLevel = "Interaction level not found.";
                    }

                    interactionsData += "Drug: " + drugName + "\n" +
                                        "Interaction Level: " + interactionLevel + "\n";
                    
                   try {
                    WebElement morebtn=driver.findElement(By.xpath("(//div[@class='style__description___2l7Ow'])["+ (i+1) +"]//span"));
                    
                    if(morebtn.isDisplayed()) {
                        morebtn.click();
                    }}catch (Exception dd) {
                    	System.out.println("No More Button available");
                    	
					}
                
                    // Scope to the drug section
                 WebElement description=driver.findElement(By.xpath("(//div[@class='style__description___2l7Ow'])["+ (i+1) +"]"));
                 descriptionText = description.getText();
                    // Handle "More" button (if available)
                   
                        Thread.sleep(1000); // Increased wait for expansion
                   
                  
                    interactionsData += "Description: " + descriptionText + "\n";
                    interactionsData += "=====================================\n";

                    Thread.sleep(1000); // Delay between iterations

                } catch (NoSuchElementException e) {
                    interactionsData += "Error extracting interaction data for element at index " + i + ": " + e.getMessage() + "\n";
                    System.out.println("Error at index " + i + ": " + e.getMessage());
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }

        return interactionsData;
    }

    
    public static String scrapeFaq(WebDriver driver) {
        String faqsData = "";

        // Use WebDriverWait with Duration (Selenium 4.x)
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        // Locate the FAQ section
        WebElement faqSection = driver.findElement(By.xpath("//div[@id='faq']"));

        try {
            // Scroll to the FAQ section to ensure all elements are visible
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", faqSection);

            // Check if there is a "Show more" button, and click it if present and clickable
            boolean showMoreExists = true;
            while (showMoreExists) {
                try {
                    WebElement showMoreButton = driver.findElement(By.xpath("//div[@id='faq']//div[contains(@class, 'Faqs__toggle___2u7v1 Faqs__more___2iZTr')]"));
                    if (showMoreButton.isDisplayed() && showMoreButton.isEnabled()) {
                        showMoreButton.click();  // Click "Show more" to reveal more FAQs
                        // Wait for the new FAQ items to load after clicking "Show more"
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(@class, 'Faqs__tile___1B58W')]")));
                        Thread.sleep(1000); // Add a slight delay to ensure all elements load
                    } else {
                        showMoreExists = false; // No more "Show more" button found, exit the loop
                    }
                } catch (NoSuchElementException e) {
                    showMoreExists = false; // No "Show more" button found, exit the loop
                }
            }

            // Scrape all the FAQ questions and answers
            List<WebElement> faqItems = driver.findElements(By.xpath("//div[@id='faq']//div[@class='DrugPane__content___3-yrB']//h3"));

            // Print the number of FAQ items found
            System.out.println("Total FAQ items found: " + faqItems.size());

        //    for (WebElement faqItem : faqItems) {
            	for(int j=0;j<faqItems.size();j++) {
                try {
                    // Extract the question and answer
                    String question = driver.findElement(By.xpath("(.//h3[contains(@class, 'Faqs__ques___1iPB9')])["+j+"]")).getText();
                    String answer = driver.findElement(By.xpath("(.//div[contains(@class, 'Faqs__ans___1uuIW')])["+j+"]")).getText();

                    // Concatenate the question and answer to the result string
                    faqsData += "Question: " + question + "\n"
                            + "Answer: " + answer + "\n"
                            + "=====================================\n";
                    
                    System.out.println("Question: " + question);
                    System.out.println("Answer: " + answer);
                    System.out.println("=====================================");


                } catch (NoSuchElementException e) {
                    System.out.println("Error extracting FAQ data: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            return "Error scraping FAQs.";
        }

        return faqsData;
    }
    public static void scrollToBottom(WebDriver driver) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        long lastHeight = (long) js.executeScript("return document.body.scrollHeight");

        while (true) {
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
            try {
                Thread.sleep(2000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            long newHeight = (long) js.executeScript("return document.body.scrollHeight");
            if (newHeight == lastHeight) {
                break;
            }
            lastHeight = newHeight;
        }
    }
    
    
}
    
