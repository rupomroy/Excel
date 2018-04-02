package ExcelRead;
public class Browser {

	public static void main(String[] args) {
		String a = DataHandler.getDataFromExcel("D:\\Excel\\Rupom.xlsx", "Sheet1" , 0, 0);
		System.out.println(a);
		String b=DataHandler.getDataFromExcel("D:\\Excel\\Rupom.xlsx", "Sheet2" , 0, 0);
		System.out.println(b);
	}

}
