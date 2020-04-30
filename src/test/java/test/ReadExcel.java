package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		File file = new File("C:\\Users\\koira\\eclipse-workspace\\ExcelDriven\\excel\\WriteExcel.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet0 = workbook.getSheetAt(0);
		// to get individual cell value
		/*
		 * XSSFRow row0 = sheet0.getRow(0); XSSFCell cell1 = row0.getCell(0);
		 * System.out.println(cell1);
		 */
		//
		for (Row row : sheet0) {
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue() + "\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + "\t");
					break;
				case BLANK:
					System.out.print("--Blank Cell---" + "\t");
					break;
				}

			}
			System.out.println();
		}
		fis.close();
	}
}
