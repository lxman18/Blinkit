package jewellery;

import java.io.FileOutputStream;
import java.lang.Thread;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;



public class flipkartjewels {

    public static void main(String[] args) throws InterruptedException {
    	
    	

        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless=new");
        options.addArguments("--disable-gpu");
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.managed_default_content_settings.images", 2);
        prefs.put("profile.managed_default_content_settings.stylesheets", 2);
        prefs.put("profile.managed_default_content_settings.fonts", 2);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        
        String baseUrl = "https://www.flipkart.com/q/tissue-paper";//560001
        
        // Set up Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Page Number");
         headerRow.createCell(1).setCellValue("BrandCategory");
         headerRow.createCell(2).setCellValue("Location");
        headerRow.createCell(3).setCellValue("URL");
        headerRow.createCell(4).setCellValue("Name");
        headerRow.createCell(5).setCellValue("MRP");
        headerRow.createCell(6).setCellValue("SP");
        headerRow.createCell(7).setCellValue("UOM");
        headerRow.createCell(8).setCellValue("Availability");
        headerRow.createCell(9).setCellValue("Offer");
        headerRow.createCell(10).setCellValue("Set of");
        headerRow.createCell(11).setCellValue("Sales Package");
        headerRow.createCell(12).setCellValue("Brand");
        headerRow.createCell(13).setCellValue("Model Name");
        headerRow.createCell(14).setCellValue("Total No of Pieces");
        headerRow.createCell(15).setCellValue("Maximum Shelf Life");
        headerRow.createCell(16).setCellValue("Ideal For");
        headerRow.createCell(17).setCellValue("Composition");
        headerRow.createCell(18).setCellValue("Skin Type");
        headerRow.createCell(19).setCellValue("Type");
        headerRow.createCell(20).setCellValue("Country of Origin");
        headerRow.createCell(21).setCellValue("Net Quantity");
        headerRow.createCell(22).setCellValue("Manufacturing Process");
        headerRow.createCell(23).setCellValue("Organic");
        headerRow.createCell(24).setCellValue("Container Type");
        headerRow.createCell(25).setCellValue("Key Features");
        headerRow.createCell(26).setCellValue("Other Features");
        
       
        // Shutdown hook to save progress on interruption
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        final String interruptFilePath = ".\\Output\\Flipkart_interrupt_" + timestamp + ".xlsx";
        final XSSFWorkbook finalWorkbook = workbook; // For lambda access
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            try (FileOutputStream out = new FileOutputStream(interruptFilePath)) {
                finalWorkbook.write(out);
                System.out.println("Interrupted save: " + interruptFilePath);
            } catch (Exception e) {
                System.out.println("Error saving interrupt: " + e.getMessage());
            } finally {
                try {
                    finalWorkbook.close();
                } catch (Exception ignored) {}
            }
        }));
        
        int page = 1;
        boolean hasNextPage = true;
        int rowNum = 1; // Start after header
        String pincode = "110015"; // Hardcoded pincode, change as needed
        
        while (hasNextPage) {
            String currentUrl = (page == 1) ? baseUrl : baseUrl + "&page=" + page;
            
            Thread.sleep(3000);
            driver.get(currentUrl);
            driver.manage().window().maximize();
            
            // Wait for products to load
            wait.until(ExpectedConditions.numberOfElementsToBeMoreThan(By.xpath("//div[@class='slAVV4']"), 0));
            
            List<WebElement> products = driver.findElements(By.xpath("//div[@class='slAVV4']"));
            
            if (products.isEmpty()) {
                hasNextPage = false;
                break;
            }
            
            String currentHandle = driver.getWindowHandle();
            
            for (WebElement product : products) {
                try {
                    product.click();
                    
                    // Wait for new tab to open
                    Set<String> handles = new HashSet<>(driver.getWindowHandles());
                    String newHandle = null;
                    while (newHandle == null) {
                        Set<String> updatedHandles = driver.getWindowHandles();
                        for (String h : updatedHandles) {
                            if (!h.equals(currentHandle)) {
                                newHandle = h;
                                break;
                            }
                        }
                        Thread.sleep(5000); // Small wait to allow tab to open
                    }
                    Thread.sleep(3000);
                    driver.switchTo().window(newHandle);
                    
//                    // Set pincode for availability
//                    try {
//                        try {
//							WebElement pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//							    By.xpath("//input[contains(@placeholder, 'Enter pincode') or contains(@placeholder, 'Delivery Location')]")));
//							pincodeInput.click();
//
//							pincodeInput.clear();
//							pincodeInput.sendKeys(pincode);
//						} catch (Exception e) {
//							WebElement pincodeInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//								    By.xpath("//div[@class='JqZtEs']")));
//								pincodeInput.clear();
//								pincodeInput.click();
//								pincodeInput.sendKeys(pincode);
//						}
//                        
//                        // Click check button
//                        WebElement checkButton = driver.findElement(By.xpath("//span[contains(text(), 'Check')] | //button[contains(text(), 'Check')]"));
//                        checkButton.click();
//                        
//                        // Wait for availability update
//                        Thread.sleep(3000);
//                    } catch (Exception e) {
//                        System.out.println("Pincode setting skipped: " + e.getMessage());
//                    }
//                    
                    // Get URL
                    String productUrl = driver.getCurrentUrl();
                    
                    // Wait for product details to load
                    wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("h1")));
                    
                    int headercount = 1;
                    try {
						WebElement ReadMore = driver.findElement(By.xpath("//button[@class='QqFHMw _4FgsLt']"));
						ReadMore.click();
					} catch (Exception e) {
						WebElement ReadMore = driver.findElement(By.xpath("//div//button[@class='QqFHMw _4FgsLt']"));
						ReadMore.click();
					}
                    
                    String Brand1 = "";
                    try {
                        WebElement Brand = driver.findElement(By.xpath("//span[@class='mEh187']"));
                        Brand1 = Brand.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Brand = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[3]/div[1]/div[2]/div[2]/div/div[1]/h1/span[1]"));
                            Brand1 = Brand.getText();
                        } catch (NoSuchElementException ex) {
                        	Brand1 = "NA";
                        }
                    }
                    // Scrape details
                    String newName = "";
                    try {
                        WebElement nameElement = driver.findElement(By.xpath("//span[@class='VU-ZEz']"));
                        newName = nameElement.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement nameElement = driver.findElement(By.xpath("//h1[@class='_6EBuvT']//span[@class='VU-ZEz']"));
                            newName = nameElement.getText();
                        } catch (NoSuchElementException ex) {
                            newName = "NA";
                        }
                    }
                    System.out.println("Product Name: " + newName);
                    
                    String spValue = "NA";
                    String originalSp1 = "";
                    String mrpValue = "NA";
                    String originalMrp1 = "";
                    try {
                        WebElement sp = driver.findElement(By.xpath("//div[@class='yRaY8j A6+E6v']"));
                        originalSp1 = sp.getText();
                        spValue = originalSp1.replace("₹", "").replace(",", "");
                    } catch (NoSuchElementException s) {
                        try {
                            WebElement sp = driver.findElement(By.xpath("//div[@class='hl05eU']//div[@class='yRaY8j A6+E6v']"));
                            originalSp1 = sp.getText();
                            spValue = originalSp1.replace("₹", "").replace(",", "");
                        } catch (Exception t) {
                        	 spValue=mrpValue;
                        }
                    }
                    System.out.println("Selling Price: " + spValue);
                    
                   
                    try {
                        WebElement mrp = driver.findElement(By.xpath("//div[@class='Nx9bqj CxhGGd']"));
                        originalMrp1 = mrp.getText();
                        mrpValue = originalMrp1.replace("₹", "").replace(",", "");
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement mrp = driver.findElement(By.xpath("//div[@class='hl05eU']//div[@class='Nx9bqj CxhGGd']"));
                            originalMrp1 = mrp.getText();
                            mrpValue = originalMrp1.replace("₹", "").replace(",", "");
                        } catch (Exception t) {
                            mrpValue = spValue;
                        }
                    }
                    System.out.println("MRP: " + mrpValue);
                    
                    String offerValue = "NA";
                    if (!mrpValue.equals(spValue) && !"NA".equals(mrpValue) && !"NA".equals(spValue)) {
                        try {
                            WebElement offer = driver.findElement(By.xpath("//*[@id=\"container\"]/div/div[3]/div[1]/div[2]/div[2]/div/div[3]/div[1]/div/div[3]"));
                            String originalOffer = offer.getText();
                            offerValue = originalOffer.replace("off", "Off");
                        } catch (Exception e) {
                            try {
                                WebElement offer = driver.findElement(By.xpath("//div[@class='UkUFwK WW8yVX dB67CR']//span"));
                                String originalOffer = offer.getText();
                                offerValue = originalOffer.replace("off", "Off");
                            } catch (Exception h) {
                                offerValue = "NA";
                            }
                        }
                    }
                    System.out.println("Offer: " + offerValue);
                    
                    //UOM (Weight)
                    String uomValue = "NA";
                    try {
                        WebElement uom = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//li[@class='aJWdJI dpZEpc']//a)[1]")));
                        uomValue = uom.getText().trim();
                    } catch (Exception e) {
                        // Alternative XPath if needed
                        try {
                            WebElement uom = driver.findElement(By.xpath("//*[@id=\"swatch-0-weight\"]/a"));
                            uomValue = uom.getText().trim();
                        } catch (Exception ex) {
                            uomValue = "NA";
                        }
                    }
                    System.out.println("UOM (Weight): " + uomValue);
                    
                    // Availability
                    int result = 1;
                    String[] textsToCheck = { "Currently Unavailable", "Currently out of stock for", "Sold Out", "NOTIFY ME" };
                    String pageSource = driver.getPageSource();
                    boolean outOfStock = false;
                    for (String text : textsToCheck) {
                        if (pageSource.toLowerCase().contains(text.toLowerCase())) {
                            outOfStock = true;
                            break;
                        }
                    }
                    if (outOfStock) {
                        result = 0;
                    } else {
                        try {
                            WebElement atc = null;
                            try {
                                atc = driver.findElement(By.xpath("//button[@class='QqFHMw vslbG+ In9uk2']"));
                            } catch (Exception eeea) {
                                try {
                                    atc = driver.findElement(By.xpath("//button[@class='QqFHMw vslbG+ _3Yl67G _7Pd1Fp']"));
                                } catch (Exception ignored) {
                                    // atc remains null
                                }
                            }
                            if (atc != null && atc.isDisplayed() && atc.isEnabled()) {
                                result = 1;
                            } else {
                                result = 0;
                            }
                        } catch (Exception eee) {
                            result = 0;
                        }
                    }
                    
                    String availability = String.valueOf(result);
                    System.out.println("Availability: " + availability);
       //Set Of
                    String SetOf = "";
                    try {
                        WebElement SetO = driver.findElement(By.xpath("(//td[contains(.,'Set of')]/following::td//li)[1]"));
                        SetOf = SetO.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement SetO = driver.findElement(By.xpath("(//td[contains(.,'Set of')]/following::td//ul//li)[1]"));
                            SetOf = SetO.getText();
                        } catch (NoSuchElementException ex) {
                        	SetOf = "NA";
                        }
                    }
                    System.out.println("SetOf: " + SetOf);
//SalesPackage
                    String SalesPackage = "";
                    try {
                        WebElement SalesPack = driver.findElement(By.xpath("(//td[contains(.,'Sales Package')]/following::td//li)[1]"));
                        SalesPackage = SalesPack.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement SalesPack = driver.findElement(By.xpath("(//td[contains(.,'Sales Package')]/following::td//ul//li)[1]"));
                            SalesPackage = SalesPack.getText();
                        } catch (NoSuchElementException ex) {
                        	SalesPackage = "NA";
                        }
                    }                    
                    System.out.println("SalesPackage: " + SalesPackage);
//Brand
                    String Brand = "";
                    try {
                        WebElement Bran = driver.findElement(By.xpath("(//td[contains(.,'Brand')]/following::td//li)[1]"));
                        Brand = Bran.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Bran = driver.findElement(By.xpath("(//td[contains(.,'Brand')]/following::td//ul//li)[1]"));
                            Brand = Bran.getText();
                        } catch (NoSuchElementException ex) {
                        	Brand = "NA";
                        }
                    }                    
                    System.out.println("Brand: " + Brand);
                    
  //Model Name
                    String ModelName = "";
                    try {
                        WebElement ModelNa = driver.findElement(By.xpath("(//td[contains(.,'Model Name')]/following::td//li)[1]"));
                        ModelName = ModelNa.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement ModelNa = driver.findElement(By.xpath("(//td[contains(.,'Model Name')]/following::td//ul//li)[1]"));
                            ModelName = ModelNa.getText();
                        } catch (NoSuchElementException ex) {
                        	ModelName = "NA";
                        }
                    }                    
                    System.out.println("ModelName: " + ModelName);
                                       
 //Total No of Pieces
                    String TotalPieces = "";
                    try {
                        WebElement TotalPi = driver.findElement(By.xpath("(//td[contains(.,'Total No of Pieces')]/following::td//li)[1]"));
                        TotalPieces = TotalPi.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement TotalPi = driver.findElement(By.xpath("(//td[contains(.,'Total No of Pieces')]/following::td//ul//li)[1]"));
                            TotalPieces = TotalPi.getText();
                        } catch (NoSuchElementException ex) {
                        	TotalPieces = "NA";
                        }
                    }                    
                    System.out.println("TotalPieces: " + TotalPieces);      
                    
//Maximum Shelf Life
                    String MaximumShelfLife = "";
                    try {
                        WebElement MaximumShelf = driver.findElement(By.xpath("(//td[contains(.,'Maximum Shelf')]/following::td//li)[1]"));
                        MaximumShelfLife = MaximumShelf.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement MaximumShelf = driver.findElement(By.xpath("(//td[contains(.,'Maximum Shelf')]/following::td//ul//li)[1]"));
                            MaximumShelfLife = MaximumShelf.getText();
                        } catch (NoSuchElementException ex) {
                        	MaximumShelfLife = "NA";
                        }
                    }                    
                    System.out.println("MaximumShelfLife: " + MaximumShelfLife);          
                    
//Ideal For
                    String IdealFor = "";
                    try {
                        WebElement IdealF = driver.findElement(By.xpath("(//td[contains(.,'Ideal For')]/following::td//li)[1]"));
                        IdealFor = IdealF.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement IdealF = driver.findElement(By.xpath("(//td[contains(.,'Ideal For')]/following::td//ul//li)[1]"));
                            IdealFor = IdealF.getText();
                        } catch (NoSuchElementException ex) {
                        	IdealFor = "NA";
                        }
                    }                    
                    System.out.println("IdealFor: " + IdealFor);                      
  //Composition
                    String Composition = "";
                    try {
                        WebElement Compositi = driver.findElement(By.xpath("(//td[contains(.,'Composition')]/following::td//li)[1]"));
                        Composition = Compositi.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Compositi = driver.findElement(By.xpath("(//td[contains(.,'Composition')]/following::td//ul//li)[1]"));
                            Composition = Compositi.getText();
                        } catch (NoSuchElementException ex) {
                        	Composition = "NA";
                        }
                    }                    
                    System.out.println("Composition: " + Composition);                           
                    
      //Skin Type
                    String SkinType = "";
                    try {
                        WebElement SkinTy = driver.findElement(By.xpath("(//td[contains(.,'Skin Type')]/following::td//li)[1]"));
                        SkinType = SkinTy.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement SkinTy = driver.findElement(By.xpath("(//td[contains(.,'Skin Type')]/following::td//ul//li)[1]"));
                            SkinType = SkinTy.getText();
                        } catch (NoSuchElementException ex) {
                        	SkinType = "NA";
                        }
                    }                    
                    System.out.println("SkinType: " + SkinType);    
                    
    //Type
                    String Typeeee = "";
                    try {
                        WebElement Typee = driver.findElement(By.xpath("(//td[contains(.,'Type')]/following::td//li)[2]"));
                        Typeeee = Typee.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Typee = driver.findElement(By.xpath("(//td[contains(.,'Type')]/following::td//ul//li)[2]"));
                            Typeeee = Typee.getText();
                        } catch (NoSuchElementException ex) {
                        	Typeeee = "NA";
                        }
                    }                    
                    System.out.println("Typeeee: " + Typeeee);               
                    
         //Country of Origin
                    String CountryofOrigin = "";
                    try {
                        WebElement CountryofOr = driver.findElement(By.xpath("(//td[contains(.,'Country of')]/following::td//li)[1]"));
                        CountryofOrigin = CountryofOr.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement CountryofOr = driver.findElement(By.xpath("(//td[contains(.,'Country of')]/following::td//ul//li)[1]"));
                            CountryofOrigin = CountryofOr.getText();
                        } catch (NoSuchElementException ex) {
                        	CountryofOrigin = "NA";
                        }
                    }                    
                    System.out.println("CountryofOrigin: " + CountryofOrigin);                   
                    
 //Net Quantity
                    String NetQuantity = "";
                    try {
                        WebElement NetQuant = driver.findElement(By.xpath("(//td[contains(.,'Net Quantity')]/following::td//li)[1]"));
                        NetQuantity = NetQuant.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement NetQuant = driver.findElement(By.xpath("(//td[contains(.,'Net Quantity')]/following::td//ul//li)[1]"));
                            NetQuantity = NetQuant.getText();
                        } catch (NoSuchElementException ex) {
                        	NetQuantity = "NA";
                        }
                    }                    
                    System.out.println("NetQuantity: " + NetQuantity);              
        //ManufacturingProcess
                    String ManufacturingProcess = "";
                    try {
                        WebElement ManufacturingP = driver.findElement(By.xpath("(//td[contains(.,'Manufacturing Process')]/following::td//li)[1]"));
                        ManufacturingProcess = ManufacturingP.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement ManufacturingP = driver.findElement(By.xpath("(//td[contains(.,'Manufacturing Process')]/following::td//ul//li)[1]"));
                            ManufacturingProcess = ManufacturingP.getText();
                        } catch (NoSuchElementException ex) {
                        	ManufacturingProcess = "NA";
                        }
                    }                    
                    System.out.println("ManufacturingProcess: " + ManufacturingProcess);                
                    
      //Organic
                    String Organic = "";
                    try {
                        WebElement Organ = driver.findElement(By.xpath("(//td[contains(.,'Organic')]/following::td//li)[1]"));
                        Organic = Organ.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement Organ = driver.findElement(By.xpath("(//td[contains(.,'Organic')]/following::td//ul//li)[1]"));
                            Organic = Organ.getText();
                        } catch (NoSuchElementException ex) {
                        	Organic = "NA";
                        }
                    }                    
                    System.out.println("Organic: " + Organic);                    
     //Container Type
                    String ContainerType = "";
                    try {
                        WebElement ContainerT = driver.findElement(By.xpath("(//td[contains(.,'Container Ty')]/following::td//li)[1]"));
                        ContainerType = ContainerT.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement ContainerT = driver.findElement(By.xpath("(//td[contains(.,'Container Ty')]/following::td//ul//li)[1]"));
                            ContainerType = ContainerT.getText();
                        } catch (NoSuchElementException ex) {
                        	ContainerType = "NA";
                        }
                    }                    
                    System.out.println("ContainerType: " + ContainerType);         
                    
//KeyFeatures
                    String KeyFeatures = "";
                    try {
                        WebElement KeyFeatu = driver.findElement(By.xpath("(//td[contains(.,'Key Features')]/following::td//li)[1]"));
                        KeyFeatures = KeyFeatu.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement KeyFeatu = driver.findElement(By.xpath("(//td[contains(.,'Key Features')]/following::td//ul//li)[1]"));
                            KeyFeatures = KeyFeatu.getText();
                        } catch (NoSuchElementException ex) {
                        	KeyFeatures = "NA";
                        }
                    }                    
                    System.out.println("KeyFeatures: " + KeyFeatures);  
  //Other Features
                    String OtherFeatures = "";
                    try {
                        WebElement OtherFeatu = driver.findElement(By.xpath("(//td[contains(.,'Other Feat')]/following::td//li)[1]"));
                        OtherFeatures = OtherFeatu.getText();
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement OtherFeatu = driver.findElement(By.xpath("(//td[contains(.,'Other Feat')]/following::td//ul//li)[1]"));
                            OtherFeatures = OtherFeatu.getText();
                        } catch (NoSuchElementException ex) {
                        	OtherFeatures = "NA";
                        }
                    }                    
                    System.out.println("OtherFeatures: " + OtherFeatures);  
                    
                    
                    
                    try {
						System.out.println("Scraped Url"+productUrl);
					} catch (Exception esd) {
						System.out.println("Failed Url"+productUrl);
					}
                    
                    System.out.println("===================================== product count :"+ (headercount++) +"=====================================================");
                    // Write to Excel
                    Row dataRow = sheet.createRow(rowNum++);
                    dataRow.createCell(0).setCellValue(page);
                    dataRow.createCell(1).setCellValue(Brand1);
                    dataRow.createCell(2).setCellValue("Bangalore");
                    dataRow.createCell(3).setCellValue(productUrl);
                    dataRow.createCell(4).setCellValue(newName);
                    dataRow.createCell(5).setCellValue(mrpValue);
                    dataRow.createCell(6).setCellValue(spValue);
                    dataRow.createCell(7).setCellValue(uomValue);
                    dataRow.createCell(8).setCellValue(availability);
                    dataRow.createCell(9).setCellValue(offerValue);
                    dataRow.createCell(10).setCellValue(SetOf);
                    dataRow.createCell(11).setCellValue(SalesPackage);
                    dataRow.createCell(12).setCellValue(Brand);
                    dataRow.createCell(13).setCellValue(ModelName);
                    dataRow.createCell(14).setCellValue(TotalPieces);
                    dataRow.createCell(15).setCellValue(MaximumShelfLife);
                    dataRow.createCell(16).setCellValue(IdealFor);
                    dataRow.createCell(17).setCellValue(Composition);
                    dataRow.createCell(18).setCellValue(SkinType);
                    dataRow.createCell(19).setCellValue(Typeeee);
                    dataRow.createCell(20).setCellValue(CountryofOrigin);
                    dataRow.createCell(21).setCellValue(NetQuantity);
                    dataRow.createCell(22).setCellValue(ManufacturingProcess);//560001
                    dataRow.createCell(23).setCellValue(Organic);
                    dataRow.createCell(24).setCellValue(ContainerType);
                    dataRow.createCell(25).setCellValue(KeyFeatures);
                    dataRow.createCell(26).setCellValue(OtherFeatures);
                    
                    
                    
                    System.out.println("---");
                    
                } catch (Exception eaq) {
                    System.out.println("Error processing product: " + eaq.getMessage());
                } finally {
                    // Close the new tab and switch back
                    try {
                        driver.close();
                        driver.switchTo().window(currentHandle);
                    } catch (Exception ex) {
                        // Ignore if already closed
                    }
                }
            }
            
            // Save after each page to preserve progress
            try {
                String progressFilePath = ".\\Output\\Flipkart_Delhi_" + page + "_" + timestamp + ".xlsx";
                try (FileOutputStream out = new FileOutputStream(progressFilePath)) {
                    workbook.write(out);
                }
                System.out.println("Progress saved after page " + page + ": " + progressFilePath);
            } catch (Exception e) {
                System.out.println("Error saving progress: " + e.getMessage());
            }
            
            page++;
            // Optional: Add max page limit, e.g., if (page > 50) { hasNextPage = false; }
        }
        
        // Final save
        String outputFilePath = ".\\Output\\Flipkart_tissue-paper_" + timestamp + ".xlsx";
        try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
            workbook.write(outFile);
            System.out.println("Final output file saved: " + outputFilePath);
        } catch (Exception e) {
            System.out.println("Error saving final Excel: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (Exception e) {
                // Ignore
            }
        }
        
        driver.quit();
    }
}