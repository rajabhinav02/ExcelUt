import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelUtility {
	
	public static void main(String[] args) throws IOException {
		System.out.println("ha");
		ArrayList<String> al = new ArrayList();
		
		FileInputStream fis = new FileInputStream("C:\\SelTestData\\TestDataone.xlsx");
		XSSFWorkbook excel = new XSSFWorkbook(fis);
		
		int sheets = excel.getNumberOfSheets();
		int k=0;
		int column = 0;
		
		for (int i=0; i<sheets; i++) {
			
			if(excel.getSheetName(i).equals("Test")) {
				System.out.println("no");
				XSSFSheet sheet=excel.getSheetAt(i);
				
				Iterator<Row> rows = sheet.iterator();
				
				Row firstrow=rows.next();
				Iterator<Cell>cs= firstrow.cellIterator();
				while(cs.hasNext()) {
					if (cs.next().getStringCellValue().equalsIgnoreCase("TestCases")) {
						column=k;
					}
					k++;
				}
				
				
				System.out.println(column);
				
				while (rows.hasNext()) {
					
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase("TC4")) {
					   Iterator<Cell>cl=r.cellIterator();
					   while (cl.hasNext()) {
					   al.add(cl.next().getStringCellValue());
					}
					}
					
				}
				System.out.println(al.get(1));
			}
			
		}
		
		
	}

}
