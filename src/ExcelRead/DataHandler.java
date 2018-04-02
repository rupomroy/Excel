package ExcelRead;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

class DataHandler{
	public static String getDataFromExcel(String Filepath, String sheetName, int row_index, int  cell_index) {
		String  data=null;
		try
		{
			File f = new File(Filepath);
			FileInputStream fis = new FileInputStream(f);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet s = wb.getSheet(sheetName);
			Row r = s.getRow(row_index);
			Cell c = r.getCell(cell_index);
			data=c.toString();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return data;
	}
}
