import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ex {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream fi = new FileInputStream("C:\\SelTestData\\TestDataone.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fi);
		int k=0;
		int column = 0;
		ArrayList<String> al = new ArrayList<String>();
		
		int nusheet=wb.getNumberOfSheets();
		
		for (int i=0; i<nusheet; i++) {
			XSSFSheet sh =wb.getSheetAt(i);
			if (sh.getSheetName().equals("Test")) {
				Iterator<Row> rows=sh.iterator();
				Row FR=rows.next();	
				Iterator<Cell> ce=FR.iterator();
				while(ce.hasNext()) {
					if (ce.next().getStringCellValue().equals("TestCases")) {
					column=k;	
					}
					k++;
				}
				while(rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equals("TC3")){
						Iterator<Cell>cl = r.iterator();
						while (cl.hasNext()) {
							/*if (cl.next().getStringCellValue().equals("TC3")) {
								cl.next();
							}else {
							al.add(cl.next().getStringCellValue());
							}*/
							
							if (!cl.next().getStringCellValue().equals("TC3")) {
								al.add(cl.next().getStringCellValue());
							}else {
								cl.next();
							}
						}
					}
					
				}
				System.out.println(al.get(0));
			}
			
		
			
			
		}
		
	}

}
