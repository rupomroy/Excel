package ExcelRead;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DataHandlerOne {
	public static String getdatafromexcel(String filepath, String sheetname, int row, int cell )
	{
		String data=null;
		try
		{
			File f=new File(filepath);
			FileInputStream fis=new FileInputStream(f);
			Workbook wb=WorkbookFactory.create(fis);
			Sheet ss= wb.getSheet(sheetname);
			  Row rr = ss.getRow(row);
			Cell c1  =rr.getCell(cell);
		data	=c1.toString();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return data;
		
	}

}
