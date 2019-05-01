/*package excel.main;

// clean package assembly:single
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

*//**
 * Hello world!
 *
 *//*
public class oldoldmain {
	public static void main(String[] args) throws InvalidFormatException {
		try {
			System.out.println("File Name/Destination is: " + args[0]);
			String appropriatePath = args[0].replace("\\", "\\\\");
			String newPathName = appropriatePath.substring(0,appropriatePath.lastIndexOf("\\\\")+2);
			System.out.println("Approprating Name/Destination for JVM to understand: " + appropriatePath);
			
			// C:\\Users\\faraz\\Desktop\\For Faraz.xlsx
			File f = new File(appropriatePath);
			final FileInputStream file = new FileInputStream(f);
			System.out.println("File Description: " + f.getAbsolutePath());
			final XSSFWorkbook workbook = new XSSFWorkbook(file);

			System.out.println("File Opened Successfully: ");

			System.out.println("Getting Months... ");
			String[] names = new String[workbook.getNumberOfSheets()];
			for (int i = 0; i < names.length; i++) {
				names[i] = workbook.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2);
			}
			System.out.println("Months in the sheet are: ");
			for (int i = 0; i < names.length; i++) {
				System.out.print(names[i] + " ");
			}
			System.out.println();

			Set<String> set = new HashSet<String>();
			for (int i = 0; i < names.length; i++) {
				set.add(names[i]);
			}
			System.out.println("Unique Months are : ");
			for (String s : set) {
				System.out.print(s + " ");
			}
			System.out.println();

			XSSFWorkbook workbooknew = null;
			System.out.println("Saving New Files to: " + newPathName);

			for (String s : set) {
				workbooknew = new XSSFWorkbook(new File(appropriatePath));
				for (int i = workbooknew.getNumberOfSheets() - 1; i >= 0; i--) {
					if (!s.contentEquals(workbooknew.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2))) {
						workbooknew.removeSheetAt(i);
					}
				}
				String newFileName = monthName(Integer.parseInt(s)) + "_" + Calendar.getInstance().getTimeInMillis();
				FileOutputStream out = new FileOutputStream(
						new File(newPathName + newFileName + ".xlsx"));
				workbooknew.write(out);
				out.close();
				System.out.println("Created File " + newFileName + " Successfully...");
			}
			file.close();
			workbook.close();
			System.out.println("Aborting Program.");
			
			long startTime = System.nanoTime();
			long endTime = System.nanoTime();
			long elapsedTime = endTime - startTime;
			double seconds = (double)elapsedTime / 1000000000.0;
			//System.out.println("Took "+ TimeUnit.SECONDS.convert(elapsedTime, TimeUnit.NANOSECONDS) + " seconds"); 
			System.out.println("Took "+ seconds + " seconds"); //0.949094623 - 0.956024469 - 0.931489724

		} catch (FileNotFoundException e) {
			System.out.println("File not found");
		} catch (IOException e) {
			e.printStackTrace();
		} catch (java.lang.ArrayIndexOutOfBoundsException e) {
			System.out.println("No File Name given to open the file");
		}
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
}
*/