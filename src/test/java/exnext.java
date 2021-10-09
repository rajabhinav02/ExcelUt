import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exnext {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream fj = new FileInputStream("C:\\SelTestData\\TestDataone.xlsx");
		XSSFWorkbook wb =new XSSFWorkbook(fj);
		ArrayList<String>al = new ArrayList();
		int totsheets=wb.getNumberOfSheets();
		int k=0;
		int column = 0;
		
		for (int i=0; i<totsheets; i++) {
			XSSFSheet sheet=wb.getSheetAt(i);
			
			if (sheet.getSheetName().equalsIgnoreCase("Test")){
				
				Iterator<Row>rows=sheet.rowIterator();
				Row fr =rows.next();
				
				Iterator<Cell>cells=fr.iterator();
				
				while(cells.hasNext()) {
					if(cells.next().getStringCellValue().equals("TestCases")) {
						column=k;
					}
					k++;
				}
				
				while(rows.hasNext()) {
					Row fg = rows.next();
					if (fg.getCell(column).getStringCellValue().equals("TC2")){
						Iterator<Cell> cg = fg.iterator();
						while(cg.hasNext()) {
							Cell c = cg.next();
							if(c.getCellType()==CellType.STRING) {
								al.add(c.getStringCellValue());
							}
							else {
								String h=NumberToTextConverter.toText(c.getNumericCellValue());
								al.add(h);
								
							}
						}
					}
				
					
				}
			break;	
			}
			
		}
		System.out.println(al.get(1));		
	}

}
