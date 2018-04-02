package WriteData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Writeprogram {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File f=new File("D:\\Excel\\Rupomselenoumone.xlsx");
		FileInputStream f1=new FileInputStream(f);
		Workbook wb=WorkbookFactory.create(f1);
		Sheet s1=wb.getSheet("Sheet1");
		Row r1=s1.createRow(0);
		Cell c1=r1.createCell(0);
		c1.setCellValue("wipro");
		FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		Sheet s2=wb.getSheet("Sheet2");
		Row r2=s2.createRow(1);
		Cell c2=r2.createCell(1);
		      c2.setCellValue("tcs");
		      FileOutputStream fo1=new FileOutputStream(f);
		                     wb.write(fo1);
		                     Sheet s3=wb.getSheet("Sheet3");
		             		Row r3=s3.createRow(1);
		             		Cell c3=r3.createCell(1);
		             		      c3.setCellValue("infosys");
		             		      FileOutputStream fo2=new FileOutputStream(f);
		             		                     wb.write(fo2);
		
	
		



	
		


	}

}
