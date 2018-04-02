package WriteData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DataHandlersTwo {
	public static void storedatatoexcel(String filepath, String sheetname, int row, int cell, String data )
	{
		
		try
		{
			File f=new File(filepath);
			FileInputStream fis=new FileInputStream(f);
			Workbook wb=WorkbookFactory.create(fis);
			Sheet ss= wb.getSheet(sheetname);
			  Row rr = ss.createRow(row);
			Cell c1  =rr.createCell(cell);
		          c1.setCellValue(data);
		         FileOutputStream fos=new FileOutputStream(f);
		         wb.write(fos);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
		
	}

}
