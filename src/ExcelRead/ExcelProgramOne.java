package ExcelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProgramOne {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File f=new File("D:\\Excel\\Rupom.xlsx");
		FileInputStream fis=new FileInputStream(f);
		  Workbook wb =WorkbookFactory.create(fis);
		                 Sheet s1 =wb.getSheet("Sheet1");
		                      Row r1  = s1.getRow(0);
		                  Cell c1   = r1.getCell(0);
		            String s2 =c1.toString();
		            System.out.println(s2);
		            Sheet s3 =wb.getSheet("Sheet2");
                    Row r3  = s3.getRow(0);
                Cell c3   = r3.getCell(0);
          String s4 =c3.toString();
          System.out.println(s4);
		            
		

	}

}
