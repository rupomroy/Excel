package ExcelRead;

public class Browserone {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
	String a=DataHandlerOne.getdatafromexcel("D:\\Excel\\Rupom.xlsx", "Sheet1", 0, 0);
	System.out.println(a);
	String b=DataHandlerOne.getdatafromexcel("D:\\Excel\\Rupom.xlsx", "Sheet2", 0,0);
	System.out.println(b);
	}

}
