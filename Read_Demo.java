package data_driven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Demo {
	
	
	
  public static void main(String[] args) throws IOException {
	
	 File f=new File("C:\\Users\\Kavin\\eclipse-workspace\\Maven_11\\Data.xlsx ") ;
	  
	  FileInputStream fi = new  FileInputStream(f);
	  
	  Workbook w =new XSSFWorkbook(fi);
	  
	  Sheet sheetAt = w.getSheetAt(0);
	  
	  Row row = sheetAt.getRow(5);
	  
	  Cell cell = row.getCell(1);
	  
	  CellType cellType = cell.getCellType();
	  
	  
	  if (cellType.equals(cellType.STRING)) {
		  
		System.out.println(cell.getStringCellValue());
		
	} else if (cellType.equals(cellType.NUMERIC)) {
		
		//System.out.println((cell.getNumericCellValue()));
		double numericCellValue = cell.getNumericCellValue();

		int value = (int)numericCellValue;
		System.out.println(value);
	}
	  
	  
	  
	  
}
}
