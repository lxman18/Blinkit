package downloadvideo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.time.Duration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class videodownload {
    public static void main(String[] args) throws Exception {

        String downloadPath = "C:\\Download";

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadPath);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);
        

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        FileInputStream fis = new FileInputStream(new File(".\\input-data\\CSE1 Video.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(1);
        Iterator<Row> rowIterator = sheet.iterator();
        
        // Create a new Workbook to store video names
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Downloaded Videos");
        Row headerRow = resultSheet.createRow(0);
        headerRow.createCell(0).setCellValue("Video Name");
        int rowCount = 1;
        

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            String url = row.getCell(0).getStringCellValue();

            System.out.println("Processing: " + url);

            driver.get("https://y2mate.lol/en158/");
            Thread.sleep(2000);
            driver.switchTo().defaultContent();

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

            WebElement inputTextBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@class='form-control input-lg']")));
            inputTextBox.clear();
            inputTextBox.sendKeys(url);

            WebElement downloadButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='btn-submit']")));
            downloadButton.click();

            Thread.sleep(4000);

            try {
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("document.querySelector('div[style*=\"z-index: 2147483647\"]').remove();");
            } catch (Exception ignored) {}

            driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='widgetv2Api']")));

            WebElement videoTab = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link'])[2]")));
            videoTab.click();

            WebElement videoDownload = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//tr[td[contains(text(), '360p (.mp4)')]]//button[@class='btn btn-success y2link-download custom']")));
            videoDownload.click();

            Thread.sleep(2000);

            String parentWindow = driver.getWindowHandle();
            for (String windowHandle : driver.getWindowHandles()) {
                if (!windowHandle.equals(parentWindow)) {
                    driver.switchTo().window(windowHandle);
                    driver.close();
                }
            }
            driver.switchTo().window(parentWindow);

            List<WebElement> iframes = driver.findElements(By.tagName("iframe"));
            if (!iframes.isEmpty()) {
                driver.switchTo().frame(0);
            }

            Thread.sleep(5000);

            WebElement popupDownloadClick = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='form-group has-success has-feedback']//a//span")));
            popupDownloadClick.click();
            
            Thread.sleep(7000);

            String parentWindow2 = driver.getWindowHandle();
            for (String windowHandle : driver.getWindowHandles()) {
                if (!windowHandle.equals(parentWindow2)) {
                    driver.switchTo().window(windowHandle);
                    driver.close();
                }
            }
            driver.switchTo().window(parentWindow2);

            boolean crDownloadStarted = false;
            int checkCount = 0;
            while (checkCount < 10) {
                File[] tempFiles = new File(downloadPath).listFiles((dir, name) -> name.endsWith(".crdownload"));
                if (tempFiles != null && tempFiles.length > 0) {
                    crDownloadStarted = true;
                    break;
                }
                Thread.sleep(1000);
                checkCount++;
            }

            if (!crDownloadStarted) {
                System.out.println("Download did not start for: " + url);
                continue;
            }

            waitForDownloadComplete(downloadPath, 240);

            File latestFile = getLatestDownloadedFile(downloadPath);
            if (latestFile != null && latestFile.exists()) {
                System.out.println("Downloaded: " + latestFile.getAbsolutePath());
                
                // Write the video name to the result Excel sheet
                Row resultRow = resultSheet.createRow(rowCount++);
                resultRow.createCell(0).setCellValue(url);
                resultRow.createCell(1).setCellValue(latestFile.getName());
                
            } else {
                System.out.println("No downloaded file found for: " + url);
            }
        }
        
     // Save the result Excel file with video names
        FileOutputStream fos = new FileOutputStream(".\\Output\\DownloadedVideos.xlsx");
        resultWorkbook.write(fos);
        fos.close();

        driver.quit();
        workbook.close();
        fis.close();
    }

    public static void waitForDownloadComplete(String downloadDir, int timeoutSeconds) throws InterruptedException {
        File dir = new File(downloadDir);
        int waited = 0;
        long lastSize = -1;
        File downloadedFile = null;

        while (waited < timeoutSeconds) {
            boolean downloading = false;
            File[] files = dir.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.getName().endsWith(".crdownload")) {
                        downloading = true;
                        break;
                    }
                }
            }

            if (!downloading) {
                File latestFile = getLatestDownloadedFile(downloadDir);
                if (latestFile != null && latestFile.getName().endsWith(".mp4")) {
                    long currentSize = latestFile.length();
                    if (currentSize == lastSize) {
                        return;
                    } else {
                        lastSize = currentSize;
                        downloadedFile = latestFile;
                    }
                }
            }

            Thread.sleep(1000);
            waited++;
        }

        System.out.println("Download timeout reached. File may be incomplete: " +
                (downloadedFile != null ? downloadedFile.getAbsolutePath() : "No file found"));
    }

    public static File getLatestDownloadedFile(String dirPath) {
        File dir = new File(dirPath);
        File[] files = dir.listFiles((d, name) -> name.endsWith(".mp4"));
        if (files == null || files.length == 0) return null;

        File lastModifiedFile = files[0];
        for (File file : files) {
            if (file.lastModified() > lastModifiedFile.lastModified()) {
                lastModifiedFile = file;
            }
        }
        return lastModifiedFile;
    }
}
