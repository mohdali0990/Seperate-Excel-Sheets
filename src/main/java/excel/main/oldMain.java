/*package excel.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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
public class oldMain 
{
    public static void main( String[] args ) throws InvalidFormatException
    {
try {
	
			FileInputStream file = new FileInputStream(new File("C:\\Users\\faraz\\Desktop\\For Faraz.xlsx"));
			
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			int n = workbook.getNumberOfSheets();
			System.out.println("Number of sheets are: " + n);
			//Get first sheet from the workbook
			XSSFSheet sheet1 = workbook.getSheetAt(0);
			XSSFSheet sheet2 = workbook.getSheetAt(1);
			XSSFSheet sheet3 = workbook.getSheetAt(2);
			
			String [] names = new String[n];
			for (int i = 0; i < names.length; i++) {
				names[i] = workbook.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2);
			}
			System.out.println("Names of sheet are: ");
			for (int i = 0; i < names.length; i++) {
				System.out.print(names[i] + " ");
			}
			System.out.println();
			String []str = new String[n];
			for (int i = 0; i < str.length; i++) {
				str[i] = names[i].split("[\\(\\)]")[1];
			}
			
			System.out.println("Date are: ");
			
			for (int i = 0; i < str.length; i++) {
				System.out.print(str[i] + " ");
			}
			
			String [] months = new String[n];
			for (int i = 0; i < months.length; i++) {
				months[i] = str[i].substring(0, 2);
			}
			System.out.println();
			System.out.println("Months are: ");
			for (int i = 0; i < months.length; i++) {
				System.out.println(months[i]);
			}
			
			
			Set<String> set = new HashSet<String>();
			for (int i = 0; i < names.length; i++) {
				set.add(names[i]);
			}
			System.out.println("Distinct Months Numbers are : ");
			for (String s : set) {
				System.out.println(s);
			}
			
			//List<String> list = new ArrayList<String>(set);
			XSSFWorkbook workbooknew = null;
			List<XSSFWorkbook> bookList = new ArrayList<XSSFWorkbook>();
			
			for (String s : set) {
				
				workbooknew = new XSSFWorkbook(new File("C:\\Users\\faraz\\Desktop\\For Faraz.xlsx"));
				
				for (int i = workbooknew.getNumberOfSheets()-1; i >= 0; i--) {
					
					if (!s.contentEquals(workbooknew.getSheetName(i).split("[\\(\\)]")[1].substring(0, 2))){
					String sheetName = workbooknew.getSheetName(i);
					System.out.println(sheetName);
					workbooknew.removeSheetAt(i);
					}
				}
				
				
				FileOutputStream out = 
						new FileOutputStream(new File("C:\\Users\\faraz\\Desktop\\"+monthName(Integer.parseInt(s))+".xlsx"));
					workbooknew.write(out);
					out.close();					
					bookList.add(workbooknew);
			}
			
			
			
			
			
			
			for (int i = 0; i < bookList.size(); i++) {
				System.out.println(bookList.get(i).hashCode());
			}		
			XSSFSheet sheets; 
			List<XSSFSheet> sheetList = new ArrayList<XSSFSheet>();
			for (int i = 0; i < n; i++) {
				sheets = workbook.getSheetAt(i);
				sheetList.add(sheets);
			}
			
			for (int i = 0; i < sheetList.size(); i++) {
				System.out.println(sheetList.get(i));
			}
			
			for (XSSFWorkbook book : bookList) {
				for (XSSFSheet sheet : sheetList) {
					
				}
				
				
			}
			
			file.close();
			workbook.close();
			FileOutputStream out = 
				new FileOutputStream(new File("C:\\Users\\faraz\\Google Drive\\taha bhai stuff\\test.xlsx"));
			workbook.write(out);
			out.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
    public static String monthName(int n){
    	if(n==1){
    		return "January";
    	} else if(n==2){
    		return "February";
    	} else if(n==3){
    		return "March";
    	} else if(n==4){
    		return "April";
    	} else if(n==5){
    		return "May";
    	} else if(n==6){
    		return "June";
    	} else if(n==7){
    		return "July";
    	} else if(n==8){
    		return "August";
    	} else if(n==9){
    		return "September";
    	} else if(n==10){
    		return "October";
    	} else if(n==11){
    		return "November";
    	} else if(n==12){
    		return "December";
    	} else {
    		return "could not read month";
    	}
    	
    }
}
*/