package CommonUtility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class Screenshot {

	String stringCellValue;

	public void takeScreenshot(WebDriver driver) throws IOException{
//			String filePath =  ".\\input-data\\Input Data2.xlsx";
//			FileInputStream fileInputStream = new FileInputStream(filePath);
//			Workbook workbook = new XSSFWorkbook(fileInputStream);
//			org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
//			int totalRow = sheet.getLastRowNum() + 1;
//			int columnIndexToRead = 1;
//			for (int i = 1; i <= 1 ; i++) {
//				for (int j = 1; j <= totalRow; j++) {
//						Row row = sheet.getRow(j);
//						Iterator<Cell> cellIterator = row.cellIterator();
//						Cell next = cellIterator.next();
//							Cell cell = row.getCell(i);
//							if (cell!=null) {
//								CellType cellType = cell.getCellType();
//								String stringCellValue = cell.getStringCellValue();
//								System.out.println(cell.getStringCellValue());
//								break;
////							}
//							
//						}
//						
//				}
//							java.lang.String stringCellValue = cell.getStringCellValue();
//							String screenshotFilePath = "Screenshot_"+ stringCellValue + currentDate() + ".png";
//						
								// Take a screenshot
								TakesScreenshot screenshot = (TakesScreenshot) driver;
								File srcFile = screenshot.getScreenshotAs(OutputType.FILE);
							
								// Define the destination file path and name
								String destinationPath = "C:\\Users\\kiruthiga.5785\\git\\POCbeauty\\Project_01\\Screenshot\\" + "Screenshot_" + currentDate() + ".png" ;
							
								// Save the screenshot to the destination file

								try {
									FileUtils.copyFile(srcFile, new File(destinationPath));
									
								} catch (IOException e) {
									e.printStackTrace();
								}
								System.out.println("Screenshot saved to: " + destinationPath);
						}
						
				
//			}
								
								

	public static String currentDate() {

		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy_HH_mm_ss");
		LocalDateTime currentDate = LocalDateTime.now();
		String date = dtf.format(currentDate);
		return " " + date + "";
	}

//	public static void main(String[] args) {
//		Screenshot s = new Screenshot();
//		s.takeScreenshot(null);
//		
}
