package test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel 
{
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet0 = workbook.createSheet("data1");
		//XSSFRow row0 = sheet0.createRow(0);
		
		/*
		 * Cell cell1 =row0.createCell(0); 
		 * Cell cell2=row0.createCell(1);
		 * 
		 * cell1.setCellValue("Firstname"); 
		 * cell2.setCellValue("Lastname");
		 */
		
		for(int rows=0;rows<10;rows++)
		{
			XSSFRow row = sheet0.createRow(rows);
			
			for(int cols=0;cols<10;cols++)
			{
				Cell cell = row.createCell(cols);
				cell.setCellValue((int)(Math.random()*100));
			}
		}
		
		
		File file = new File("C:\\Users\\koira\\eclipse-workspace\\ExcelDriven\\excel\\WriteExcel.xlsx");
		FileOutputStream fo = new FileOutputStream(file);
		workbook.write(fo);
		
		fo.close();
		System.out.println("File is created");
		
		
	}

}
