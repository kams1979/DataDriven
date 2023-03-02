package data_driven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Demo {
	
	public static void main(String[] args) throws IOException {
		
		 File f=new File("C:\\Users\\Kavin\\eclipse-workspace\\Maven_11\\Data.xlsx ");
		 
		 FileInputStream fi = new FileInputStream(f);

		  Workbook w =new XSSFWorkbook(fi);
		  
		  w.createSheet("sheet5").createRow(0).createCell(0).setCellValue("Name");
		  
		  w.getSheet("sheet5").createRow(1).createCell(0).setCellValue("Phonenumber");
		  
		  w.getSheet("sheet5").createRow(2).createCell(0).setCellValue("Place");
		  
		  w.getSheet("sheet5").createRow(3).createCell(0).setCellValue("Designation");
		  
		  w.getSheet("sheet5").createRow(4).createCell(0).setCellValue("Salary");
		  
		  w.getSheet("sheet5").getRow(0).createCell(1).setCellValue("kams");
		  
		  w.getSheet("sheet5").getRow(1).createCell(1).setCellValue("99128879");
		  
		  w.getSheet("sheet5").getRow(2).createCell(1).setCellValue("chennai");
		  
		  w.getSheet("sheet5").getRow(3).createCell(1).setCellValue("Software");
		  
		  w.getSheet("sheet5").getRow(4).createCell(1).setCellValue("50000");
		  
		  FileOutputStream fo=new  FileOutputStream(f);
		  
		  w.write(fo);
		  
		  w.close();
		  
		  
		  
	}

	
	
}
