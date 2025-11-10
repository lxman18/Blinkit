package Zepto_Pharma;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class zeptoraghi {

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();

        int count = 0;
        String spValue = "";
        String offerValue = "NA";
        String newName = null;
        String mrpValue = null;
        String originalMrp1 = " ";
        String originalMrp2 = " ";
        String originalMrp3 = " ";
        String originalSp1 = " ";
        String originalSp2 = " ";
        String NewAvailability1 = " ";
        String BrandName;

        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\Zepto_highlights_input.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Zepto_tissues");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            List<String> InputPlatformID = new ArrayList<>(), InputCity = new ArrayList<>(), Pincode = new ArrayList<>(), uRL = new ArrayList<>(), categores = new ArrayList<>();

            // Extract URLs from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);

                if (i == 0) {
                    continue;
                }

                Cell InputPlatformIDCell = row.getCell(0);
                Cell inputCityCell = row.getCell(1);
                Cell pinCodeCell = row.getCell(2);
                Cell urlCell = row.getCell(3);
                Cell categorescell = row.getCell(4);

                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String id = (InputPlatformIDCell != null && InputPlatformIDCell.getCellType() == CellType.STRING) ? InputPlatformIDCell.getStringCellValue() : "";
                    String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING) ? inputCityCell.getStringCellValue() : "";
                    String pin = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";
                    String category = (categorescell != null && categorescell.getCellType() == CellType.STRING) ? categorescell.getStringCellValue() : "";
                    String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";

                    InputPlatformID.add(id);
                    InputCity.add(city);
                    Pincode.add(pin);
                    uRL.add(url);
                    categores.add(category);
                    
                    Pincode.add(locationSet);
                }
            }

            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            Row headerRow = resultsSheet.createRow(0);

            // Get current date for header
            String timestamp = new SimpleDateFormat("dd-MM-yyyy").format(new Date());

            // Set header values
          //  Row headerRow = resultsSheet.createRow(rowIndex++);
            headerRow.createCell(0).setCellValue("id");
            headerRow.createCell(1).setCellValue("city");
            headerRow.createCell(2).setCellValue("pin");
            headerRow.createCell(3).setCellValue("url");
            headerRow.createCell(4).setCellValue("category");
            headerRow.createCell(5).setCellValue("newName");
            headerRow.createCell(6).setCellValue("UomValue");
            headerRow.createCell(7).setCellValue("mrpValue");
            headerRow.createCell(8).setCellValue("spValue");
            headerRow.createCell(9).setCellValue("NewAvailability1");
            headerRow.createCell(10).setCellValue("BrandName");
            headerRow.createCell(11).setCellValue("product_ty");
            headerRow.createCell(12).setCellValue("model");
            headerRow.createCell(13).setCellValue("Shape");
            headerRow.createCell(14).setCellValue("Pattern");
            headerRow.createCell(15).setCellValue("Design_type");
            headerRow.createCell(16).setCellValue("color_type");
            headerRow.createCell(17).setCellValue("Brush_type");
            headerRow.createCell(18).setCellValue("Material_type");
            headerRow.createCell(19).setCellValue("set");
            headerRow.createCell(20).setCellValue("Bristle_type");
            headerRow.createCell(21).setCellValue("Gender");
            headerRow.createCell(22).setCellValue("Keyfeatures_type");
            headerRow.createCell(23).setCellValue("closure_type");
            headerRow.createCell(24).setCellValue("gem_type");
            headerRow.createCell(25).setCellValue("size");
            headerRow.createCell(26).setCellValue("ideal_type");
            headerRow.createCell(27).setCellValue("Theme");
            headerRow.createCell(28).setCellValue("PackagingType");
            
            headerRow.createCell(29).setCellValue("item_included");
            headerRow.createCell(30).setCellValue("ProductDimesions");
            
            headerRow.createCell(31).setCellValue("unit");

            int rowIndex = 1;
            int headercount = 0;
            String currentPin = null;

            for (int i = 0; i < uRL.size(); i++) {
                String id = InputPlatformID.get(i);
                String city = InputCity.get(i);
                String pin = Pincode.get(i);
                String url = uRL.get(i);
                String category = categores.get(i);                
                String locationSet = Pincode.get(i);

                try {
                    if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                        // Set "NA" values in all columns
                        Row resultRow = resultsSheet.createRow(rowIndex++);
                        resultRow.createCell(0).setCellValue(id);
                        resultRow.createCell(1).setCellValue(city);
                        resultRow.createCell(2).setCellValue(pin);
                        resultRow.createCell(3).setCellValue(url);
                        resultRow.createCell(4).setCellValue(category);
                        resultRow.createCell(5).setCellValue("NA");
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
                        
                        System.out.println("Skipped processing for URL: " + url);
                        continue;
                    }

                    driver.get(url);
                    Thread.sleep(5000);
                    driver.manage().window().maximize();

                    // Location setting
                    if (currentPin == null || !currentPin.equals(locationSet)) {
                        Thread.sleep(9000);

                        for (int k = 0; k < 100; k++) {
                            try {
                                Thread.sleep(3000);
                                WebElement location = driver.findElement(By.xpath("//button[@aria-label='Select Location']"));
                                location.click();
                                break;
                            } catch (Exception r) {
                                try {
                                    Thread.sleep(4000);
                                    WebElement location = driver.findElement(By.xpath("/html/body/div[1]/div/header/div/div[2]/button"));
                                    location.click();
                                    break;
                                } catch (NoSuchElementException u) {
                                    try {
                                        Thread.sleep(4000);
                                        WebElement location = driver.findElement(By.xpath("/html/body/div[2]/div/div/div/div[2]/div/button[1]/div/p"));
                                        location.click();
                                        Thread.sleep(3000);
                                        break;
                                    } catch (Exception tr) {
                                        Thread.sleep(4000);
                                        WebElement location = driver.findElement(By.xpath("/html/body/div[1]/div/header/div/div[1]/button"));
                                        location.click();
                                        Thread.sleep(3000);
                                        break;
                                    }
                                }
                            }
                        }

                        String tempPinNumber = "";
                        for (int j = 0; j < 20; j++) {
                            try {
                                Thread.sleep(3000);
                                driver.findElement(By.xpath("//input[@placeholder='Search a new address']")).sendKeys(Keys.ENTER);
                                Thread.sleep(3000);
                                driver.findElement(By.xpath("//input[@placeholder='Search a new address']")).clear();
                                Thread.sleep(3000);
                                System.out.println("print the crt pin number" + locationSet);
                                String crtPin = locationSet;
                                driver.findElement(By.xpath("//input[@placeholder='Search a new address']")).sendKeys(crtPin);
                                Thread.sleep(3000);
                                try {
                                    driver.findElement(By.xpath("/html/body/div[4]/div/div/div/div/div/div/div[2]/div/div[2]/div[2]/div[1]")).click();
                                } catch (Exception e) {
                                    driver.findElement(By.xpath("(//div[@data-testid='address-search-item'])[1]")).click();
                                }
                                Thread.sleep(2000);
                                currentPin = locationSet;
                                System.out.println("=============" + currentPin + "======================");
                                Thread.sleep(3000);
                                driver.findElement(By.xpath("//button[contains(text(), 'Confirm & Continue')]")).click();
                                Thread.sleep(2000);
                                currentPin = locationSet;
                                break;
                            } catch (Exception e) {
                            }
                        }
                    }

                    Thread.sleep(2000);

   // Extract product name
                    try {
                        WebElement nameElement = driver.findElement(By.xpath("//h1[@class='cp62rX c9OiKy c7CsPX']"));
                        newName = nameElement.getText();
                        System.out.println(newName);
                    } catch (NoSuchElementException e) {
                        try {
							WebElement nameElement = driver.findElement(By.xpath("//div[contains(@class,'u-flex u-items-center')]//h1"));
							newName = nameElement.getText();
							System.out.println(newName);
						} catch (Exception e1) {
							newName="NA";
						}
                    }
                    System.out.println("headercount = " + headercount);
                    headercount++;

 // Extract selling price (SP)
                    Thread.sleep(2000);
                    try {
                        WebElement sp = driver.findElement(By.xpath("//span[@class='TvBIc']"));
                        originalSp1 = sp.getText();
                        spValue = originalSp1.replace("₹", "");
                        System.out.println(spValue);
                    } catch (Exception e) {
                        try {
                            WebElement sp = driver.findElement(By.xpath("//p//span[@class='TvBIc']"));
                            originalSp2 = sp.getText();
                            spValue = originalSp2.replace("₹", "");
                            System.out.println(spValue);
                        } catch (Exception exx) {
                            spValue = "NA";
                        }
                    }

// Extract MRP
                    Thread.sleep(2000);
                    try {
                        WebElement mrp = driver.findElement(By.xpath("//span[@class='_IKg9 dUCkZ']"));
                        originalMrp1 = mrp.getText();
                        mrpValue = originalMrp1.replace("₹", "");
                        System.out.println(mrpValue);
                    } catch (NoSuchElementException e) {
                        try {
                            WebElement mrp = driver.findElement(By.xpath("//p//span[@class='_IKg9 dUCkZ']"));
                            originalMrp2 = mrp.getText();
                            mrpValue = originalMrp2.replace("₹", "");
                            System.out.println(mrpValue);
                        } catch (Exception ex) {
                            try {
                                WebElement mrp = driver.findElement(By.xpath("//*[@id=\\\"product-features-wrapper\\\"]/div[1]/div/div[3]/div[1]/div[2]/p/span[2]"));
                                originalMrp3 = mrp.getText();
                                mrpValue = originalMrp3.replace("₹", "");
                                System.out.println(mrpValue);
                            } catch (Exception exx) {
                                mrpValue = spValue;
                                System.out.println(mrpValue);
                            }
                        }
                    }

       // Extract UOM
                    Thread.sleep(2000);
                    String Uom;
                    String UomValue;
                    try {
                        WebElement uom1 = driver.findElement(By.xpath("//span[@class='font-bold']"));
                        Uom = uom1.getText();
                        UomValue = Uom;
                    } catch (Exception e) {
                        WebElement uom1 = driver.findElement(By.xpath("(//div[@class='BoqfC']//span)[2]"));
                        Uom = uom1.getText();
                        UomValue = Uom;
                    }
                    System.out.println(UomValue);

//  //Extract item included
//                   String Item;
//                    String ItemCount;
//                    try {
//                        WebElement item = driver.findElement(By.xpath("//div[@class='u-flex u-items-start KjTQZ'][8]//p"));
//                        Item = item.getText();
//                        ItemCount = Item;
//                    } catch (Exception e) {
//                        WebElement item = driver.findElement(By.xpath("(//div[@class='w-1/2 break-words'])[8]"));
//                        Item = item.getText();
//                        ItemCount = Item;
//                    }
//                    System.out.println(ItemCount);

// Extract brand
                    try {
                        WebElement brand = driver.findElement(By.xpath("(//h3[contains(.,'brand')]/following::div[@class='w-1/2 break-words'])[1]"));
                        BrandName = brand.getText();
                    } catch (Exception e) {
                        WebElement brand = driver.findElement(By.xpath("(//div[@class='u-flex u-items-start KjTQZ']//p)[1]"));
                        BrandName = brand.getText();
                    }
                    System.out.println(BrandName);

  // Check availability
                    int result = 1;
                    try {
                        WebElement cart = driver.findElement(By.xpath("(//button[contains(text(), 'Add To Cart')])[2]"));
                        if (cart.isEnabled()) {
                            result = 1;
                        } else {
                            result = 0;
                        }
                    } catch (Exception ae) {
                        WebElement notify = driver.findElement(By.xpath("(//span[text()='Notify Me'])[2]"));
                        if (notify.isDisplayed()) {
                            result = 0;
                        }
                    }
                    System.out.println(result);
                    NewAvailability1 = String.valueOf(result);

// Extract product type
                    String product_ty = "";
                    try {
                        WebElement product_type = driver.findElement(By.xpath("(//h3[contains(text(),'product ')]/following::div)[1]//p"));
                        product_ty = product_type.getText().trim();
                        System.out.println(product_ty);
                    } catch (Exception e) {
                        
                            product_ty = "NA";
                         
                    }
//model name 
                    String model = "";
                    try {
                        WebElement mode = driver.findElement(By.xpath("(//h3[contains(text(),'model')]/following::div)[1]//p"));
                        model = mode.getText().trim();
                        System.out.println(model);
                    } catch (Exception e) {
                        
                    	model = "NA";
                          
                    }
                    
                    
 //Shape
                    String Shape = "";
                    try {
                        WebElement shap = driver.findElement(By.xpath("(//h3[contains(text(),'shap')]/following::div)[1]//p"));
                        Shape = shap.getText().trim();
                        System.out.println(Shape);
                    } catch (Exception e) {
                        
                        	Shape = "NA";
                        
                    }
                    
  //pattern
                    String Pattern = "";
                    try {
                        WebElement Patte = driver.findElement(By.xpath("(//h3[contains(text(),'patter')]/following::div)[1]//p"));
                        Pattern = Patte.getText().trim();
                        System.out.println(Pattern);
                    } catch (Exception e) {
                        try {
                            WebElement Patte = driver.findElement(By.xpath("(//h3[contains(text(),'patter')]/following::div)[1]"));
                            Pattern = Patte.getText().trim();
                            System.out.println(Pattern);
                        } catch (Exception ex) {
                        	Pattern = "NA";
                            System.out.println(Pattern);
                        }
                    }
 //  design type
                    String Design_type = "";
                    try {
                        WebElement desi_type = driver.findElement(By.xpath("(//h3[contains(text(),'design')]/following::div)[1]//p"));
                        Design_type = desi_type.getText().trim();
                        System.out.println(Design_type);
                    } catch (Exception e) {
                        try {
                            WebElement desi_type = driver.findElement(By.xpath("(//h3[contains(text(),'design')]/following::div)[1]"));
                            Design_type = desi_type.getText().trim();
                            System.out.println(Design_type);
                        } catch (Exception ex) {
                            Design_type = "NA";
                            System.out.println(Design_type);
                        }
                    }
//  colour name
                    String color_type = "";
                    try {
                        WebElement colo_type = driver.findElement(By.xpath("(//h3[contains(text(),'colo')]/following::div)[1]//p"));
                        color_type = colo_type.getText().trim();
                        System.out.println(color_type);
                    } catch (Exception e) {
                        try {
                            WebElement colo_type = driver.findElement(By.xpath("(//h3[contains(text(),'colour')]/following::div)[1]"));
                            color_type = colo_type.getText().trim();
                            System.out.println(color_type);
                        } catch (Exception ex) {
                            color_type = "NA";
                            System.out.println(color_type);
                        }
                    }
                    
         // Brush_type
                    String Brush_type = "";
                    try {
                        WebElement Bru_type = driver.findElement(By.xpath("(//h3[contains(text(),'brush')]/following::div)[1]//p"));
                        Brush_type = Bru_type.getText().trim();
                        System.out.println(Brush_type);
                    } catch (Exception e) {
                        try {
                            WebElement Bru_type = driver.findElement(By.xpath("(//h3[contains(text(),'brush')]/following::div)[1]"));
                            Brush_type = Bru_type.getText().trim();
                            System.out.println(Brush_type);
                        } catch (Exception ex) {
                        	Brush_type = "NA";
                            System.out.println(Brush_type);
                        }
                    }                   

   // Extract material type
                    String Material_type = "";
                    try {
                        WebElement mate_type = driver.findElement(By.xpath("(//h3[contains(text(),'material type')]/following::div)[1]//p"));
                        Material_type = mate_type.getText().trim();
                        System.out.println(Material_type);
                    } catch (Exception e) {
                        try {
                            WebElement mate_type = driver.findElement(By.xpath("(//h3[contains(text(),'material')]/following::div)[1]"));
                            Material_type = mate_type.getText().trim();
                            System.out.println(Material_type);
                        } catch (Exception ex) {
                            Material_type = "NA";
                            System.out.println(Material_type);
                        }
                    }
       // Extract Set
                    String set = "";
                    try {
                        WebElement se = driver.findElement(By.xpath("(//h3[contains(text(),'set')]/following::div)[1]//p"));
                        set = se.getText().trim();
                        System.out.println(set);
                    } catch (Exception e) {
                        try {
                            WebElement se = driver.findElement(By.xpath("(//h3[contains(text(),'set')]/following::div)[16]"));
                            set = se.getText().trim();
                            System.out.println(set);
                        } catch (Exception ex) {
                        	set = "NA";
                            System.out.println(set);
                        }
                    }
                
         // Extract Bristle_type
                             String Bristle_type = "";
                             try {
                                 WebElement Bri_type = driver.findElement(By.xpath("(//h3[contains(text(),'bristle type')]/following::div)[1]//p"));
                                 Bristle_type = Bri_type.getText().trim();
                                 System.out.println(Bristle_type);
                             } catch (Exception e) {
                                 try {
                                     WebElement Bri_type = driver.findElement(By.xpath("(//h3[contains(text(),'bristle type')]/following::div)[1]"));
                                     Bristle_type = Bri_type.getText().trim();
                                     System.out.println(Bristle_type);
                                 } catch (Exception ex) {
                                	 Bristle_type = "NA";
                                     System.out.println(Bristle_type);
                                 }
                             }
             // Extract Gender
                             String Gender = "";
                             try {
                                 WebElement gente = driver.findElement(By.xpath("(//h3[contains(text(),'gender')]/following::div)[1]//p"));
                                 Gender = gente.getText().trim();
                                 System.out.println(Gender);
                             } catch (Exception e) {
                                 try {
                                     WebElement gente = driver.findElement(By.xpath("(//h3[contains(text(),'gender')]/following::div)[1]"));
                                     Gender = gente.getText().trim();
                                     System.out.println(Gender);
                                 } catch (Exception ex) {
                                	 Gender = "NA";
                                     System.out.println(Gender);
                                 }
                             }                            
   //  key features
                    String Keyfeatures_type = "";
                    try {
                        WebElement key_type = driver.findElement(By.xpath("(//h3[contains(text(),'key features')]/following::div)[1]"));
                        Keyfeatures_type = key_type.getText().trim();
                        System.out.println(Keyfeatures_type);
                    } catch (Exception e) {
                        try {
                            WebElement key_type = driver.findElement(By.xpath("(//h3[contains(text(),'key features')]/following::div)[1]//p"));
                            Keyfeatures_type = key_type.getText().trim();
                            System.out.println(Keyfeatures_type);
                        } catch (Exception ex) {
                            Keyfeatures_type = "NA";
                            System.out.println(Keyfeatures_type);
                        }
                    }

//  closure type
                    String closure_type = "";
                    try {
                        WebElement closu_type = driver.findElement(By.xpath("(//h3[contains(text(),'closure')]/following::div)[1]"));
                        closure_type = closu_type.getText().trim();
                        System.out.println(closure_type);
                    } catch (Exception e) {
                        try {
                            WebElement closu_type = driver.findElement(By.xpath("(//h3[contains(text(),'closure')]/following::div)[1]//p"));
                            closure_type = closu_type.getText().trim();
                            System.out.println(closure_type);
                        } catch (Exception ex) {
                            closure_type = "NA";
                            System.out.println(closure_type);
                        }
                    }

   //  gem type
                    String gem_type = "";
                    try {
                        WebElement ge_type = driver.findElement(By.xpath("(//h3[contains(text(),'gem')]/following::div)[1]"));
                        gem_type = ge_type.getText().trim();
                        System.out.println(gem_type);
                    } catch (Exception e) {
                        try {
                            WebElement ge_type = driver.findElement(By.xpath("(//h3[contains(text(),'gem')]/following::div)[1]//p"));
                            gem_type = ge_type.getText().trim();
                            System.out.println(gem_type);
                        } catch (Exception ex) {
                            gem_type = "NA";
                            System.out.println(gem_type);
                        }
                    }
// size 
                    String size = "";
                    try {
                        WebElement si = driver.findElement(By.xpath("(//h3[contains(text(),'size')]/following::div)[1]"));
                        size = si.getText().trim();
                        System.out.println(size);
                    } catch (Exception e) {
                        
                        	size = "NA";
                       
                    }
    //  ideal for
                    String ideal_type = "";
                    try {
                        WebElement idea_type = driver.findElement(By.xpath("(//h3[contains(text(),'idea')]/following::div)[1]"));
                        ideal_type = idea_type.getText().trim();
                        System.out.println(ideal_type);
                    } catch (Exception e) {
                        try {
                            WebElement idea_type = driver.findElement(By.xpath("(//h3[contains(text(),'ideal')]/following::div)[1]//p"));
                            ideal_type = idea_type.getText().trim();
                            System.out.println(ideal_type);
                        } catch (Exception ex) {
                            ideal_type = "NA";
                            System.out.println(ideal_type);
                        }
                    }
//Theme 
                    

                    String Theme = "";
                    try {
                        WebElement Them = driver.findElement(By.xpath("(//h3[contains(text(),'theme')]/following::div)[1]"));
                        Theme = Them.getText().trim();
                        System.out.println(Theme);
                    } catch (Exception e) {
                        
                        	Theme = "NA";
                         
                    }
                    
                    
  // PackagingType
                    
                    String PackagingType = "";
                    try {
                        WebElement Packagin = driver.findElement(By.xpath("(//h3[contains(text(),'packaging type')]/following::div)[1]//p"));
                        PackagingType = Packagin.getText().trim();
                        System.out.println(PackagingType);
                    } catch (Exception e) {
                        
                    	PackagingType = "NA";
                         
                    }
                    
// Weight
                    
                    String Weight = "";
                    try {
                        WebElement Weig = driver.findElement(By.xpath("(//h3[contains(text(),'weight')]/following::div)[1]//p"));
                        Weight = Weig.getText().trim();
                        System.out.println(Weight);
                    } catch (Exception e) {
                        
                    	Weight = "NA";
                         
                    }   
// length
                    
                    String length = "";
                    try {
                        WebElement leng = driver.findElement(By.xpath("(//h3[contains(text(),'length')]/following::div)[1]//p"));
                        length = leng.getText().trim();
                        System.out.println(length);
                    } catch (Exception e) {
                        
                    	length = "NA";
                         
                    }     
// Eco_friendly
                    
                    String Eco_friendly = "";
                    try {
                        WebElement Eco_frie = driver.findElement(By.xpath("(//h3[contains(text(),'eco')]/following::div)[1]//p"));
                        Eco_friendly = Eco_frie.getText().trim();
                        System.out.println(Eco_friendly);
                    } catch (Exception e) {
                        
                    	Eco_friendly = "NA";
                         
                    }   
// ply_rating
                    
                    String ply_rating = "";
                    try {
                        WebElement ply_r = driver.findElement(By.xpath("(//h3[contains(text(),'ply')]/following::div)[1]//p"));
                        ply_rating = ply_r.getText().trim();
                        System.out.println(ply_rating);
                    } catch (Exception e) {
                        
                    	ply_rating = "NA";
                         
                    }  
// Sheet_count
                    
                    String Sheet_count = "";
                    try {
                        WebElement Sheet_cot = driver.findElement(By.xpath("(//h3[contains(text(),'shee')]/following::div)[1]//p"));
                        Sheet_count = Sheet_cot.getText().trim();
                        System.out.println(Sheet_count);
                    } catch (Exception e) {
                        
                    	Sheet_count = "NA";
                         
                    }                        
                    
   //ModelName
                    
                    
                    String ModelName = "";
                    try {
                        WebElement ModelN = driver.findElement(By.xpath("(//h3[contains(text(),'model ')]/following::div)[1]//p"));
                        ModelName = ModelN.getText().trim();
                        System.out.println(ModelName);
                    } catch (Exception e) {
                        
                    	ModelName = "NA";
                         
                    }
  //  item included
                    String item_included = "";
                    try {
                        WebElement it_type = driver.findElement(By.xpath("(//h3[contains(text(),'item')]/following::div)[1]//p"));
                        item_included = it_type.getText().trim();
                        System.out.println(item_included);
                    } catch (Exception e) {
                        try {
                            WebElement it_type = driver.findElement(By.xpath("(//h3[contains(text(),'item')]/following::div)[1]//p"));
                            item_included = it_type.getText().trim();
                            System.out.println(item_included);
                        } catch (Exception ex) {
                            item_included = "NA";
                            System.out.println(item_included);
                        }
                    }
//ProductDimesions
                    
                    String ProductDimesions = "";
                    try {
                        WebElement ProductDimes = driver.findElement(By.xpath("(//h3[contains(text(),'product dimensions')]/following::div)[1]//p"));
                        ProductDimesions = ProductDimes.getText().trim();
                        System.out.println(ProductDimesions);
                    } catch (Exception e) {
                        
                    	ProductDimesions = "NA";
                         
                    }
  // Extract pack of
//                    String pack_off = "";
//                    try {
//                        WebElement pa_type = driver.findElement(By.xpath("(//h3[contains(text(),'pack')]/following::div)[1]//p"));
//                        pack_off = pa_type.getText().trim();
//                        System.out.println(pack_off);
//                    } catch (Exception e) {
//                        try {
//                            WebElement pa_type = driver.findElement(By.xpath("(//h3[contains(text(),'pack')]/following::div)[1]//p"));
//                            pack_off = pa_type.getText().trim();
//                            System.out.println(pack_off);
//                        } catch (Exception ex) {
//                            pack_off = "NA";
//                            System.out.println(pack_off);
//                        }
//                    }

   // Extract unit
                    String unit = "";
                    try {
                        WebElement un_type = driver.findElement(By.xpath("(//h3[contains(text(),'unit')]/following::div)[1]//p"));
                        unit = un_type.getText().trim();
                        System.out.println(unit);
                    } catch (Exception e) {
                        try {
                            WebElement un_type = driver.findElement(By.xpath("(//h3[contains(text(),'unit')]/following::div)[1]//p"));
                            unit = un_type.getText().trim();
                            System.out.println(unit);
                        } catch (Exception ex) {
                            unit = "NA";
                            System.out.println(unit);
                        }
                    }

                    // Write data to result row
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(pin);
                    resultRow.createCell(3).setCellValue(url);
                    resultRow.createCell(4).setCellValue(category);
                    resultRow.createCell(5).setCellValue(newName);
                    resultRow.createCell(6).setCellValue(UomValue);
                    resultRow.createCell(7).setCellValue(mrpValue);
                    resultRow.createCell(8).setCellValue(spValue);
                    resultRow.createCell(9).setCellValue(NewAvailability1);
                    resultRow.createCell(10).setCellValue(BrandName);
                    resultRow.createCell(11).setCellValue(product_ty);
                    resultRow.createCell(12).setCellValue(model);
                    resultRow.createCell(13).setCellValue(Shape);
                    resultRow.createCell(14).setCellValue(Pattern);
                    resultRow.createCell(15).setCellValue(Design_type);
                    resultRow.createCell(16).setCellValue(color_type);
                    resultRow.createCell(17).setCellValue(Brush_type);
                    resultRow.createCell(18).setCellValue(Material_type);
                    resultRow.createCell(19).setCellValue(set);
                    resultRow.createCell(20).setCellValue(Bristle_type);
                    resultRow.createCell(21).setCellValue(Gender);
                    resultRow.createCell(22).setCellValue(Keyfeatures_type);
                    resultRow.createCell(23).setCellValue(closure_type);
                    resultRow.createCell(24).setCellValue(gem_type);
                    resultRow.createCell(25).setCellValue(size);
                    resultRow.createCell(26).setCellValue(ideal_type);
                    resultRow.createCell(27).setCellValue(Theme);
                    resultRow.createCell(28).setCellValue(PackagingType);
                    resultRow.createCell(29).setCellValue(item_included);
                    resultRow.createCell(30).setCellValue(ProductDimesions);
                    resultRow.createCell(31).setCellValue(unit);
                    

                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(pin);
                    resultRow.createCell(3).setCellValue(url);
                    resultRow.createCell(4).setCellValue(category);
                    resultRow.createCell(5).setCellValue(newName);
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
                    resultRow.createCell(30).setCellValue("NA");
                    resultRow.createCell(31).setCellValue("NA");
                    
                    System.out.println("Failed to extract data for URL: " + url);
                }
            }

            // Write results to Excel file
            String outputTimestamp = new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss").format(new Date());
            FileOutputStream outFile = new FileOutputStream(".\\Output\\Zepto_Tissues_Output" + outputTimestamp + ".xlsx");
            resultsWorkbook.write(outFile);
            outFile.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                System.out.println("DoNe DoNe Scraping DoNe");
                driver.quit();
            }
        }
    }
}