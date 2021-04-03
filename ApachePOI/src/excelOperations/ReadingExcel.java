package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {
	
	/*
	 * Giving error NPE - NULL POinter Exception while using for loop
	 * With iterator we are able to get data without NPE
	 */

	public static void main(String[] args) throws IOException {

		String excelFilePath = ".\\dataFiles\\worldcities.xlsx"; 
		//place where my .xlsx file is present {.\\} --> denotes present directory
		
		//To open a file in reading mode we use FileInputStream
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream); //got the workbook to read from
		
//		XSSFSheet sheet = workbook.getSheet("Sheet1"); //getting sheet to read from Workbook
		XSSFSheet sheet = workbook.getSheetAt(0); // only need to specify the sheet number which starts with 0 itself.
		
		// Using For Loop 
		
		/*int rows = sheet.getLastRowNum(); //This will get the last row from the sheet specified above.
		int col = sheet.getRow(1).getLastCellNum(); //This will get row 1 inside which how any cols are there.
		
		for(int r=0; r<=rows; r++) // represents row in a excel 
		{
			//get the row.
			XSSFRow row = sheet.getRow(r); //0 row
			
			for(int c=0; c<=col; c++) // represents columns in a excel 
			{
				// this loop will get all the columns corresponding to 0 row
				//get the col
				XSSFCell cell = row.getCell(c);
				try {
				switch(cell.getCellType()) 
				{
				// It will get the type of the data stored in particular cell as all the cell might not be of String type

				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				case BLANK: break;
				case ERROR: break;
				case FORMULA: break;
				case _NONE: break;
				default: break;
				}
				System.out.print(" || ");
				} catch(NullPointerException e) {
					e.printStackTrace();
				}
			}
			System.out.println();
		}*/
		
		// Using Iterator :
		
		Iterator<Row> iterator = sheet.iterator(); // row iterator
		
		while(iterator.hasNext()) {
			XSSFRow row = (XSSFRow) iterator.next(); // getting row
			
			Iterator<Cell> cellIterator = row.cellIterator(); // cell iterator
			
			while(cellIterator.hasNext()) {
				XSSFCell cell = (XSSFCell) cellIterator.next(); // getting cell
				switch(cell.getCellType()) // checking the datatype of data contained in cell 
				{
				// It will get the type of the data stored in particular cell as all the cell might not be of String type

				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				case BLANK: break;
				case ERROR: break;
				case FORMULA: break;
				case _NONE: break;
				default: break;
				}
				System.out.print(" || ");
				} 
			System.out.println();
			}
			
		}
		
	}


