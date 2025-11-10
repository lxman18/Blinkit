package Youtubecode;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;
import java.util.NoSuchElementException;

public class dummy3 {
    public static void main(String[] args) throws Exception {

        // ===== Set download preferences =====
        String downloadFilepath = "C:\\Users\\AktharJohn.6880\\Music"; // your folder path
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadFilepath);
        prefs.put("download.prompt_for_download", false);
        prefs.put("download.directory_upgrade", true);
        prefs.put("safebrowsing.enabled", true);

        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", prefs);

        // ===== Use your existing Chrome profile (so Gmail is already logged in) =====
        options.addArguments("user-data-dir=C:\\Users\\AktharJohn.6880\\AppData\\Local\\Google\\Chrome\\User Data");
        options.addArguments("--profile-directory=Default"); 
        // 🔼 If your Gmail login is in "Profile 1" or "Profile 2", change above accordingly.

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        driver.manage().window().maximize();

        // ===== Read input Excel file =====
        String excelPath = "C:\\Users\\AktharJohn.6880\\Pictures\\Book1.xlsx";  // <-- your input file path
        FileInputStream fis = new FileInputStream(new File(excelPath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int url_count = 1;

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {   // start from row 1 (skip header)
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell cell = row.getCell(0); // assuming URLs are in column A
            if (cell == null) continue;

            String videoUrl = cell.getStringCellValue().trim();
            if (videoUrl.isEmpty()) continue;

            try {
                System.out.println("🎬 Processing: " + videoUrl);

                // ===== Open ssyoutube with video link =====
                driver.get("https://ssyoutube.com");

                WebElement search = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("//input[@placeholder='Paste your video link here']")));
                search.clear();
                search.sendKeys(videoUrl);
                Thread.sleep(2000);
                search.sendKeys(Keys.ESCAPE);
                Thread.sleep(2000);
                search.sendKeys(Keys.ENTER);

                // Wait for download options to appear
                Thread.sleep(2000);
                wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//div[@class='col-12 col-lg-8 order-2']")));
                Thread.sleep(3000);

                // ===== Try clicking best available quality =====
                WebElement downloadBtn = null;
                Thread.sleep(3000);
                try {
                    try {
                        Thread.sleep(1500);
                        downloadBtn = driver.findElement(By.xpath("//a[@id='download-mp4-360-audio']"));
                        downloadBtn.click();

                    } catch (NoSuchElementException e1) {
                        Thread.sleep(1500);
                        downloadBtn = driver.findElement(By.xpath("//a[@id='download-mp4-144-no-audio']"));
                        downloadBtn.click();
                    }
                } catch (Exception e) {
                    Thread.sleep(1500);
                    downloadBtn = driver.findElement(By.xpath("//a[@id='download-mp4-240-audio']"));
                    downloadBtn.click();
                }

                // ===== Verify download actually starts =====
                File dir = new File(downloadFilepath);
                int beforeCount = dir.listFiles().length;

                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", downloadBtn);

                boolean downloaded = false;
                for (int j = 0; j < 30; j++) { // wait up to 30 seconds
                    int afterCount = dir.listFiles().length;
                    if (afterCount > beforeCount) {
                        downloaded = true;
                        break;
                    }
                    Thread.sleep(1000);
                }

                if (downloaded) {
                    System.out.println("✅ Download started for: " + videoUrl);
                } else {
                    System.out.println("⚠️ Download may have failed: " + videoUrl);
                }

                // ===== Handle popup tabs (ads) =====
                String mainWindow = driver.getWindowHandle();
                Set<String> handles = driver.getWindowHandles();
                for (String handle : handles) {
                    if (!handle.equals(mainWindow)) {
                        driver.switchTo().window(handle);
                        driver.close();
                    }
                }
                driver.switchTo().window(mainWindow);

                // Random wait between downloads (3–6 sec)
                Thread.sleep(3000 + new Random().nextInt(3000));

            } catch (Exception e) {
                System.out.println("❌ Failed for URL: " + videoUrl);
                e.printStackTrace();
            }
            System.out.println("========================= Download_video_count: " + url_count++ + " ================================");
        }

        workbook.close();
        driver.quit();
        System.out.println("🏁 All downloads completed.");
    }
}
