package test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDrivern {
	//Identify Testcases coloum by scanning the entire 1st row
	//once coloumn is identified then scan entire testcase coloum to identify purcjhase testcase row
	//after you grab purchase testcase row = pull all the data of that row and feed into test
	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis = new FileInputStream("C:\\Users\\koira\\eclipse-workspace\\ExcelDriven\\excel\\DataForExcel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets =workbook.getNumberOfSheets();
		for (int i = 0; i <sheets; i++) 
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Data1"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows =sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell>ce = firstRow.cellIterator();
				int k = 0;
				int column =0;
				while(ce.hasNext())
				{
					Cell Value = ce.next();
					if(Value.getStringCellValue().equalsIgnoreCase("TestCase"))
					{
						column = k;
					}
					k++;
				}
				System.out.println(column);
				
				
				  while(rows.hasNext()) { Row r =rows.next();
				  if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
				  Iterator<Cell> cv = r.cellIterator(); while(cv.hasNext()) {
				  System.out.println(cv.next().getStringCellValue()); } } }
				 
			}
			
			
		}
	
	
		
	}

}
