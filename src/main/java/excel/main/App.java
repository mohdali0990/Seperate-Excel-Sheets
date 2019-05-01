package excel.main;

// clean package assembly:single
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.swing.JFileChooser;
import javax.swing.JFrame;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class App {
	public static void main(String[] args) throws InvalidFormatException {
		try {
			
			File selectedFile = OpenAndSelectFile();			
			
			System.out.println("File Selected is: " + selectedFile.getAbsolutePath());

			//Set<String> uniqueMonths = getUniqueMonths(selectedFile);

			String newPathName = getPathToSaveNewFiles(selectedFile); 			
			
			//PartTabsOutOfExcel(selectedFile,uniqueMonths, newPathName);
			//CopyTabs(selectedFile,uniqueMonths, newPathName);
			CopySampleExcel(selectedFile, newPathName);
			
		} catch (FileNotFoundException e) {
			System.out.println("File not found");
		} catch (NullPointerException e) {
			System.out.println("No file selected to open. Aborting program");
			e.printStackTrace();
		} catch (Exception e) {
			System.out.println("Something strange happened that I was not ready for. \nOnly .xlsx files are supported. Try again...");
			e.printStackTrace();
		} finally {
			System.out.println("Aborting Program.");
			System.exit(0);
		}

	}

	

	private static void CopySampleExcel(File selectedFile, String newPathName) throws IOException, InvalidFormatException {
		 
		
		 XSSFWorkbook workbooknew = new XSSFWorkbook(selectedFile);
			

			//String newFileName = monthName(Integer.parseInt(s)) + "_" + Calendar.getInstance().getTimeInMillis();
			FileOutputStream out = new FileOutputStream(new File(newPathName + selectedFile.getName() + Calendar.getInstance().getTimeInMillis() + ".xlsx"));
			workbooknew.write(out);
			out.close();
			System.out.println("Created File "  + " Successfully...");
				
	}



	private static File OpenAndSelectFile() {

		JFileChooser jFileChooser = new JFileChooser();
		jFileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));

		int result = jFileChooser.showOpenDialog(new JFrame());

		File selectedFile = null;
		if (result == JFileChooser.APPROVE_OPTION) {
			selectedFile = jFileChooser.getSelectedFile();
		}
		return selectedFile;
	}
	
	private static void PartTabsOutOfExcel(File selectedFile, Set<String> uniqueMonths, String newPathName) throws InvalidFormatException, IOException {
		XSSFWorkbook workbooknew = null;
		for (String s : uniqueMonths) {
			workbooknew = new XSSFWorkbook(selectedFile);
			for (int i = workbooknew.getNumberOfSheets() - 1; i >= 0; i--) {
				if (!s.contentEquals(workbooknew.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2))) {
					workbooknew.removeSheetAt(i);
				}
			}

			String newFileName = monthName(Integer.parseInt(s)) + "_" + Calendar.getInstance().getTimeInMillis();
			FileOutputStream out = new FileOutputStream(new File(newPathName + newFileName + ".xlsx"));
			workbooknew.write(out);
			out.close();
			System.out.println("Created File " + newFileName + " Successfully...");
		}
	}
	
	private static Set<String> getUniqueMonths(File selectedFile) throws IOException, InvalidFormatException {
		XSSFWorkbook workbook = new XSSFWorkbook(selectedFile);

		System.out.println("File Opened Successfully: ");

		String[] names = new String[workbook.getNumberOfSheets()];
		Set<String> set = new HashSet<String>();
		for (int i = 0; i < names.length; i++) {
			names[i] = workbook.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2);
			set.add(names[i]);
		}

		System.out.println("Unique Months are : ");
		for (String s : set) {
			System.out.print(s + " ");
		}
		System.out.println();
		workbook.close();
		return set;
	}

	private static String getPathToSaveNewFiles(File selectedFile) {
		String s = selectedFile.getAbsolutePath().substring(0, selectedFile.getAbsolutePath().lastIndexOf("\\") + 1);
		System.out.println("Saving New Files to: " + s);
		return s;
	}
	
	public static String monthName(int n) {
		if (n == 1) {
			return "January";
		} else if (n == 2) {
			return "February";
		} else if (n == 3) {
			return "March";
		} else if (n == 4) {
			return "April";
		} else if (n == 5) {
			return "May";
		} else if (n == 6) {
			return "June";
		} else if (n == 7) {
			return "July";
		} else if (n == 8) {
			return "August";
		} else if (n == 9) {
			return "September";
		} else if (n == 10) {
			return "October";
		} else if (n == 11) {
			return "November";
		} else if (n == 12) {
			return "December";
		} else {
			return "could not read month";
		}

	}
	
	@SuppressWarnings("deprecation")
	public static void CopyTabs(File selectedFile, Set<String> uniqueMonths, String newPathName) throws IOException, InvalidFormatException{
		   DataFormatter formatter = new DataFormatter();
		   Workbook wb = new XSSFWorkbook(selectedFile);		   
		   Row row;
	       Cell cell;
	       Workbook wb2 = null;
		for (String s : uniqueMonths) {
			wb2 = new XSSFWorkbook();
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				XSSFSheet sheetFromOldWB = (XSSFSheet) wb.getSheetAt(i);
				if (s.contentEquals(sheetFromOldWB.getSheetName().split("[\\(\\)]")[1].substring(0, 2))) {
					XSSFSheet sheetForNewWB = (XSSFSheet) wb2.createSheet(sheetFromOldWB.getSheetName());
					for (int rowIndex = 0; rowIndex < sheetFromOldWB.getPhysicalNumberOfRows(); rowIndex++) {
						row = sheetForNewWB.createRow(rowIndex);
						for (int colIndex = 0; colIndex < sheetFromOldWB.getRow(rowIndex)
								.getPhysicalNumberOfCells(); colIndex++) {
							cell = row.createCell(colIndex);
							String text = formatter.formatCellValue(sheetFromOldWB.getRow(rowIndex).getCell(colIndex));
							cell.setCellValue(text);							
						}
					}
					
				}
				
			}
			String newFileName = monthName(Integer.parseInt(s)) + "_" + Calendar.getInstance().getTimeInMillis();
			FileOutputStream fileOut = new FileOutputStream(new File(newPathName + newFileName + ".xlsx"));					
			wb2.write(fileOut);
			fileOut.close();
			
		}
	}
}
