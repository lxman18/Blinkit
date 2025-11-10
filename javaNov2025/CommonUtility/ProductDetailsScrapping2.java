package CommonUtility;

import java.time.Duration;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.NoSuchElementException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;

public class ProductDetailsScrapping2 {
	
	BlinkitId screenshot = new BlinkitId();

	public void waitHelper(RemoteWebDriver driver, String xpath) {

		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(Duration.ofSeconds(UserUtility.waitTime))
				.pollingEvery(Duration.ofSeconds(1)).ignoring(NoSuchElementException.class);

		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
	}

	public String getText(RemoteWebDriver driver, String xpath) {
		String text = driver.findElement(By.xpath(xpath)).getText();

		return text;
	}
	
//
//	/***
//	 * Format for Nykaa => productCodeReturn(URL, "/p/", "\\?");
//	 * 
//	 * @param url
//	 * @param separator1
//	 * @param separator2
//	 * @return - Product code from the URL
//	 */
//
//	public String productCodeReturn(String url, String separator1, String separator2) {
//
//		String productCode;
//
//		String currentUrl = url;
//
//		String[] splitUrl = currentUrl.split(separator1);
//
//		List<String> listUrl = new ArrayList<String>();
//
//		List<String> listUrl2 = new ArrayList<String>();
//
//		String[] splitUrl2 = null;
//
//		for (String string : splitUrl) {
//			listUrl.add(string);
//		}
//
//		String dummyURL = listUrl.get(1);o
//
//		splitUrl2 = dummyURL.split(separator2);
//
//		for (String string : splitUrl2) {
//			listUrl2.add(string);
//		}
//
//		productCode = listUrl2.get(0);
//
//		return productCode;
//	}

	public List<String> productDetails(RemoteWebDriver driver, HashMap<String, String> locatorsMap,
			String productCodeSplit, boolean qtySplit, String inputProductCode, String BlinkitPid, String BrandName) {

		List<String> dataList = new ArrayList<String>();

		String productCode = "NA";
		String URL = "NA";
		String brandName = "NA";
		String productName = "NA";
		String mrp = "NA";
		String sellingPrice = "NA";
		String quantity = "NA";
		String availability = "NA";
		String multiplier = "NA";
		String offer = "NA";
		String manualInterventionFlag = "NA";
		String logFile = "";
		String InputPin = "";
		
		

//		Input string URL
		
		if (!inputProductCode.equals("NA")) {
			String InputURL, brandNameXpath, productNameXpath,productNameXpath1, mrpXpath,mrpXpath1, sellingPriceXpath, sellingPriceXpath1,
					qtyXpath,qtyXpath1, availabilityXpath, availabilityXpath1,offerXpath;
			try {
				InputURL = locatorsMap.get("InputURL");

				InputURL = InputURL + inputProductCode;


				if (!InputURL.isBlank()) {
					driver.get(InputURL);
					URL = driver.getCurrentUrl();
				}
			}

			catch (Exception e) {

			}
			
			try { 
				 InputPin = locatorsMap.get("InputPin");
			}
			catch (Exception e) {
				  
			}
			
			if (locatorsMap.get("BrandName").equals("DotAndKey") || locatorsMap.get("BrandName").equals("JustHerbs") || locatorsMap.get("BrandName").equals("PepNails") || locatorsMap.get("BrandName").equals("Nykaa") 
					|| locatorsMap.get("BrandName").equals("SUGARPOP")|| locatorsMap.get("BrandName").equals("SUGAR")  ) {
				try {
					Thread.sleep(3000);
				} catch (InterruptedException e1) {
					e1.printStackTrace();
				}
				System.out.println("Sleep executed !!!");

			}
			
//			Product name 
			try {
				productNameXpath = locatorsMap.get("productNameXpath");

				if (!productNameXpath.isBlank()) {
					System.out.println("productName is executed");
					waitHelper(driver, productNameXpath);
					productName = getText(driver, productNameXpath);
				}

				try {
					brandNameXpath = locatorsMap.get("brandNameXpath");

					if (!brandNameXpath.isBlank()) {
						brandName = getText(driver, brandNameXpath);
					}
				}

				catch (Exception e) {
					

				}

				if (!productName.isBlank() && !brandName.equals("NA")) {
					productName = brandName + productName;
				}
			}

			catch (Exception e) {
				
				try {
					productNameXpath1 = locatorsMap.get("productNameXpath1");

					if (!productNameXpath1.isBlank()) {
						System.out.println("productName is executed");
						productName = getText(driver, productNameXpath1);
					}

					try {
						brandNameXpath = locatorsMap.get("brandNameXpath");

						if (!brandNameXpath.isBlank()) {
							brandName = getText(driver, brandNameXpath);
						}
					}

					catch (Exception e1) {
						

					}

					if (!productName.isBlank() && !brandName.equals("NA")) {
						productName = brandName + productName;
					}
					
				} 
				
				catch (Exception e2) {
					
					manualInterventionFlag = "YES";
					System.out.println(e.getMessage());
					logFile = logFile + "Product Name is not available /---------";
				}

			}
			
			//Screenshot
			
			try {
				screenshot.screenshot(driver, BrandName ,BlinkitPid );
			} catch (Exception e) {
				e.fillInStackTrace();
			
			}
//	
		
//			Selling price
			
			
			try {
				
				sellingPriceXpath = locatorsMap.get("sellingPriceXpath");
			
			//	sellingPriceXpath1 = locatorsMap.get("sellingPriceXpath1");
					if (!(sellingPriceXpath.isEmpty())) {
						
					/*	if(sellingPriceXpath == "//span[@class='a-price a-text-price a-size-medium apexPriceToPay'][1]") {

							System.out.println("Selling xpath =======" + sellingPriceXpath);  */
							
							sellingPrice = getText(driver, sellingPriceXpath);
						    sellingPrice = UserUtility.rupeesSplit(sellingPrice);
					//	} 
					/*	else {
							
							System.out.println("Selling 1 xpath =======" + sellingPriceXpath1);
							
							sellingPrice = getText(driver, sellingPriceXpath1);
						    sellingPrice = UserUtility.rupeesSplit(sellingPrice);
						    
						}    */
						
						if (sellingPrice.isBlank() || sellingPrice.isEmpty()) {
							throw new Exception("Selling price is empty string") ;
						}
					//	System.out.println("======  SELL PRICE IN THE WEBSITES  ==== " + sellingPrice);
					}
			
			}
			
			catch (Exception e) {
				
				try {
					
					sellingPriceXpath1 = locatorsMap.get("sellingPriceXpath1");
					
					if (!(sellingPriceXpath1.isEmpty())) {
						sellingPrice = getText(driver, sellingPriceXpath1);
						sellingPrice = UserUtility.rupeesSplit(sellingPrice);
						
						
					}

				} catch (Exception e2) {
					manualInterventionFlag = "YES";
					System.out.println(e.getMessage());
					logFile = logFile + "Selling price is not available /---------";
				}
			}
			
//			MRP price
			
			try {
				mrpXpath = locatorsMap.get("mrpXpath");
				
			//	mrpXpath1 = locatorsMap.get("mrpXpath1");

				if (!mrpXpath.isBlank()) {
					
					/*	if(mrpXpath == "//span[@class = 'a-price a-text-price a-size-base']//following::span[1]") {
						
						System.out.println("mrpXpath   ============   " + mrpXpath);  */
						
					mrp = getText(driver, mrpXpath);
					mrp = UserUtility.rupeesSplit(mrp);
				//	}
				/*	else {
						
						System.out.println("mrpXpath  1   ============   " + mrpXpath1);
						
						mrp = getText(driver, mrpXpath1);
						mrp = UserUtility.rupeesSplit(mrp);
					}   */
				}
				
				if (mrp.isBlank() || (mrp.equals("NA"))) {
					mrp = sellingPrice;
				}
			}

			catch (Exception e) {
				
				try {
					mrpXpath1 = locatorsMap.get("mrpXpath");

					if (!mrpXpath1.isBlank()) {
						mrp = getText(driver, mrpXpath1);
						mrp = UserUtility.rupeesSplit(mrp);
					}
					
					if (mrp.isBlank() || (mrp.equals("NA"))) {
						mrp = sellingPrice;
					}
				} 
				
				catch (Exception e2) {
					
					mrp = sellingPrice;
					
//					manualInterventionFlag = "YES";
//					System.out.println(e.getMessage());
//					logFile = logFile + "MRP price is not available /---------";
				}
				
				
			}
			try {
				offerXpath = locatorsMap.get("offerXpath");
				
				if (offerXpath.isBlank()) {
					offer = "NA";
					
				}else {
					offer = getText(driver, offerXpath);
				}
				
			}
			catch(Exception e){
				manualInterventionFlag = "YES";
				System.out.println(e.getMessage());
				logFile = logFile + "Offer is not available /---------";
			}
		}
		
		dataList.add(productCode);
		dataList.add(URL);
		dataList.add(productName);
		dataList.add(mrp);
		dataList.add(sellingPrice);
		dataList.add(quantity);
		dataList.add(availability);
		dataList.add(multiplier);
		dataList.add(offer);
		dataList.add(manualInterventionFlag);
		dataList.add(logFile);
		
		return dataList;
	}
	
	public boolean checkAvailability(RemoteWebDriver driver) {
		

		try {
			driver.findElement(By.xpath("//*[contains(text(),'Out Of Stock')]"));
			return true;
		} 
		catch (Exception e) {
			try {
				driver.findElement(By.xpath("//*[contains(text(),'Out of Stock')]"));
				return true;

			} catch (Exception e2) {
				try {
					driver.findElement(By.xpath("//*[contains(text(),'Out of stock')]"));
					return true;

				} 
				catch (Exception e3) {
					try {
						driver.findElement(By.xpath("//*[contains(text(),'Sold Out')]"));
						return true;

					} 
					catch (Exception e4) {
						try {
							driver.findElement(By.xpath("//*[contains(text(),'Sold out')]"));
							return true;

						} 
						catch (Exception e5) {
							try {
								driver.findElement(By.xpath("//*[contains(text(),'sold out')]"));
								return true;

							} 
							catch (Exception e6) {
								try {
									driver.findElement(By.xpath("//*[contains(text(),'Currently unavailable')]"));
									return true;

								} catch (Exception e7) {
									try {
										driver.findElement(By.xpath("//*[contains(text(),'Currently not available')]"));
										return true;

									} catch (Exception e8) {
										try {
											driver.findElement(By.xpath("//*[contains(text(),'Currently out of stock in this area.')]"));
											return true;
										}
											catch(Exception e9) {
										return false;
											}
									}
								}
							}
						}
					}
				}
			}
		}
			
	}	
	
}


