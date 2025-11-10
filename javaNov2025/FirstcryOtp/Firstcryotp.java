import java.time.Duration;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.mail.*;
import javax.mail.internet.MimeMultipart;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Firstcryotp {

    // ---------- Fetch OTP from Gmail ----------
    public static String waitForOTP(String user, String password, int maxWaitSeconds) {
        String otp = null;
        int waited = 0;
        while (otp == null && waited < maxWaitSeconds) {
            otp = getLatestOTP(user, password);
            if (otp != null) break;
            try {
                Thread.sleep(5000); // Wait 5 seconds before checking again
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
            // IMAP configuration
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

            // Get the latest 5 emails to increase chances of finding OTP
            Message[] messages = inbox.getMessages();
            if (messages.length == 0) {
                System.out.println("No emails found in inbox.");
                inbox.close(false);
                store.close();
                return null;
            }

            // Check the latest 5 emails (or fewer if not available)
            int start = Math.max(messages.length - 5, 0);
            for (int i = messages.length - 1; i >= start; i--) {
                Message message = messages[i];
                String from = message.getFrom()[0].toString();
                String subject = message.getSubject();

                // Filter for FirstCry OTP emails
                if (from.contains("firstcry") || subject.toLowerCase().contains("otp")) {
                    System.out.println("Checking email from: " + from + ", Subject: " + subject);
                    String content = getTextFromMessage(message);
                    // System.out.println("Email content: " + content); // Uncomment for debugging

                    // Extract OTP using regex (6-digit number)
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

    // Helper method to extract text from email (handles Multipart and HTML)
    private static String getTextFromMessage(Message message) throws Exception {
        if (message.isMimeType("text/plain")) {
            return message.getContent().toString();
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            return getTextFromMimeMultipart(mimeMultipart);
        } else if (message.isMimeType("text/html")) {
            String html = message.getContent().toString();
            // Basic HTML to text conversion (strip tags)
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
                // Basic HTML to text conversion
                result.append(html.replaceAll("<[^>]+>", "").replaceAll("\\s+", " ").trim());
            } else if (bodyPart.getContent() instanceof MimeMultipart) {
                result.append(getTextFromMimeMultipart((MimeMultipart) bodyPart.getContent()));
            }
        }
        return result.toString();
    }

    // ---------- Main Selenium Flow ----------
    public static void main(String[] args) {
        // Replace with your actual Gmail and App Password
        final String GMAIL_USER = "firstcry1stbaby0369@gmail.com";
        final String GMAIL_APP_PASSWORD = "dwhtcunljwesoamf"; // Ensure this is correct

        // Setup ChromeDriver with options
        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless=new"); // Uncomment for headless mode
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--start-maximized");
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {
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
//			WebElement Otppaste = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div//input[@name='notp[]']")));
//			Otppaste.click();
//			
			
            System.out.println("Initiated OTP login.");

            // Fetch OTP
            String otp = waitForOTP(GMAIL_USER, GMAIL_APP_PASSWORD, 90);
            if (otp != null) {
                System.out.println("Fetched OTP: " + otp);
                // Find all OTP input fields
                List<WebElement> otpInputs = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div//input[@name='notp[]']")));
                if (otpInputs.size() >= otp.length()) {
                    // Split OTP into individual digits and enter them
                    for (int i = 0; i < otp.length(); i++) {
                        WebElement otpField = otpInputs.get(i);
                        otpField.clear();
                        otpField.sendKeys(String.valueOf(otp.charAt(i)));
                    }
                    System.out.println("Entered OTP digits into fields.");
                } else {
                    System.out.println("❌ Not enough OTP input fields found. Expected: " + otp.length() + ", Found: " + otpInputs.size());
                }

                // Submit OTP
                WebElement submitOtp = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[.='SUBMIT']")));
                submitOtp.click();
                System.out.println("Submitted OTP and logged in.");
            } else {
                System.out.println("❌ OTP not received within 90 seconds.");
            }

        } catch (Exception e) {
            System.err.println("Error during execution: " + e.getMessage());
            e.printStackTrace();
        } finally {
            driver.quit();
            System.out.println("Browser closed.");
        }
    }
}