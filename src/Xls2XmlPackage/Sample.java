package Xls2XmlPackage;

import java.io.FileOutputStream;

public class Sample {

	/**
	 * @param args
	 */
	public static void main(String[] ar) throws Exception{
		// TODO Auto-generated method stub
		StringBuffer cells = new StringBuffer("PRODUCTCODE,PRODUCTFEATUREQUALIFIER,PRODUCTFEATUREVALUE,\r\n");
		String temp="dsfas";
		cells.append("sdfsdf "+temp);
		System.out.println(cells);
		FileOutputStream fos=new FileOutputStream("C:\\Users\\880324\\Desktop\\testResult\\abc.txt");
		fos.write(cells.toString().getBytes());
		fos.write("\n".getBytes());
		fos.close();
		
	}

}
