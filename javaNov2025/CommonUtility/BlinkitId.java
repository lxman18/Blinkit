package CommonUtility;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class BlinkitId {

	public static String cellValue;
	
	public static void screenshot (WebDriver driver,String brandName ,String productId) throws IOException {
		
//			String filePath = ".\\input-data\\Input Data2.xlsx";
//			Workbook workbook = new XSSFWorkbook(filePath);
//			Sheet sheet = workbook.getSheetAt(0);
//			int columnIndex = 1;
//			int numberOfRows = sheet.getPhysicalNumberOfRows();
//			for (int i = 1; i < numberOfRows + 1; i++) {
//				Row row = sheet.getRow(i);
//				if (row != null) {
//					Cell cell = row.getCell(columnIndex);
//					if (cell != null) {
//						CellType cellType = cell.getCellType();
//						String cellValue = cell.getStringCellValue();
//						System.out.println(cellValue);
//						
						
						try {
							TakesScreenshot screenshot = (TakesScreenshot) driver;
							File srcFile = screenshot.getScreenshotAs(OutputType.FILE);

							// Define the destination file path and name
							String destinationPath = "Screenshot\\" +brandName+"_"+productId+"_"+currentDate() + ".png";

							// Save the screenshot to the destination file

								FileUtils.copyFile(srcFile, new File(destinationPath));

							System.out.println("Screenshot saved to: " + destinationPath);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
//					
//				}
//				
//			}
//			
//		}
			

	public static String currentDate() {

		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy_HH_mm_ss");
		LocalDateTime currentDate = LocalDateTime.now();
		String date = dtf.format(currentDate);
		return " " + date + "";
	}


}
