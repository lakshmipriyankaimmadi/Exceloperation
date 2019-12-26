import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadDataFromExcel {

	public static void main(String[] args) throws Throwable {
		
		FileInputStream fi=new FileInputStream("d://dummy.xlsx");
		Workbook wb=WorkbookFactory.create(fi);
		Sheet sh= wb.getSheet("Login");
		Row r=sh.getRow(0);
		Cell c=r.getCell(0);
		int rowcount=sh.getLastRowNum();
		int colcount=r.getLastCellNum();
		System.out.println(+rowcount +" " +colcount);
		for(int i=0;i<=rowcount;i++){
			String username=sh.getRow(i).getCell(0).getStringCellValue();
			if(wb.getSheet("Login").getRow(i).getCell(1).getCellType()==Cell.CELL_TYPE_NUMERIC)
			{
				int celldata=(int)wb.getSheet("Login").getRow(i).getCell(1).getNumericCellValue();
				String password=String.valueOf(celldata);
				System.out.println(username+"  "+password);
			/*String password=sh.getRow(i).getCell(1).getStringCellValue();
			System.out.println(username+" "+password); */
					
		}
			sh.getRow(i).getCell(2).setCellValue("Pass");
			}
		fi.close();
		FileOutputStream fo=new FileOutputStream("d://dummy2.xlsx");
	    wb.write(fo);
	    fo.close();		
		wb.close();
	}

}
