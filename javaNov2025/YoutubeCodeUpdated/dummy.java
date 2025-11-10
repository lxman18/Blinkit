package Youtubecode;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class dummy {
    public static void main(String[] args) throws Exception {

        // ===== Set download preferences =====
        String downloadFilepath = "C:\\Users\\AktharJohn.6880\\Downloads"; // your folder path
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadFilepath);
        prefs.put("download.prompt_for_download", false);
        prefs.put("download.directory_upgrade", true);
        prefs.put("safebrowsing.enabled", true);

        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        // ===== Read input Excel file =====
        String excelPath = "C:\\Users\\AktharJohn.6880\\Pictures\\Book1.xlsx";  // <-- your input file path
        FileInputStream fis = new FileInputStream(new File(excelPath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {   // start from row 1 (skip header)
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cell = row.getCell(0); // assuming URLs are in column A
            if (cell == null) continue;

            String videoUrl = cell.getStringCellValue().trim();
            if (videoUrl.isEmpty()) continue;

            System.out.println("Processing: " + videoUrl);

            // ===== Open ssyoutube with video link =====
            driver.get("https://ssyoutube.com");
            Thread.sleep(2000);

            WebElement search = driver.findElement(By.xpath("//input[@placeholder='Paste your video link here']"));
            search.clear();

            // Paste URL
            search.sendKeys(videoUrl);
            Thread.sleep(1000);

            // Close dropdown suggestion
            search.sendKeys(Keys.ESCAPE);
            Thread.sleep(500);

            // Submit
            search.sendKeys(Keys.ENTER);

            // Wait for page to load download button
            Thread.sleep(4000);

            // Find & click 360p download button
            try {
				try {
					WebElement download = driver.findElement(By.xpath("//a[@id='download-mp4-360-audio']"));
					((JavascriptExecutor) driver).executeScript("arguments[0].click();", download);
					System.out.println("=========================================Movie to Catch=====================================================");
					
				} catch (Exception e) {
					WebElement download = driver.findElement(By.xpath("//a[@id='download-mp4-240-audio']"));
					((JavascriptExecutor) driver).executeScript("arguments[0].click();", download);
					System.out.println("=========================================Movie to Catch=====================================================");
				}
			} catch (Exception e) {
				WebElement download = driver.findElement(By.xpath("//a[@id='download-mp4-144-no-audio']"));
				((JavascriptExecutor) driver).executeScript("arguments[0].click();", download);
				System.out.println("=========================================Download video=====================================================");
			}

            // ===== Handle new tab =====
            String mainWindow = driver.getWindowHandle();
            for (String handle : driver.getWindowHandles()) {
                if (!handle.equals(mainWindow)) {
                    driver.switchTo().window(handle);
                    driver.close();
                }
            }
            driver.switchTo().window(mainWindow);

            // Small wait between downloads
            Thread.sleep(5000);
        }

        workbook.close();
        driver.quit();
        System.out.println("✅ All downloads completed.");
    }
}
