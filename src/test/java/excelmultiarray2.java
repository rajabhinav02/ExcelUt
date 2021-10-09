import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excelmultiarray2 {

	Workbook wb;
	public void getTestArray(String sheetname) {
		
		FileInputStream fis=null;
		
		try {
			fis = new FileInputStream("C:\\Test Data\\TestExcel.xlsx");
			try {
				Workbook wb = WorkbookFactory.create(fis);
				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		
		Sheet sheet= wb.getSheet("data");
		
		Object[][]obj = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		
		for (int i=0; i<sheet.getLastRowNum(); i++) {
			for (int j=0; j<sheet.getRow(0).getLastCellNum(); j++) {
				obj[i][j]=sheet.getRow(i+1).getCell(j).toString();
				System.out.println(obj[i][j]);
			}
		}
	}
	
}
