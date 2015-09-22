package Xls2XmlPackage;

import java.util.Scanner;

public class CsvCalling {
	public static void main(String[] ags) throws Exception{
		Scanner s=new Scanner(System.in);
		System.out.println("Enter input file path");
		String input=s.nextLine().trim(),output=s.nextLine().trim();		
		CsvConverter csv=new CsvConverter();	
		
		//csv.Apparelcsv(input,output);	
		
		
		csv.ElectronicCsv(input,output);
		//C:\Users\880324\Desktop\testResult\Product - Apparel - PCM Test Data - Bulk Upload.xls
		//C:\Users\880324\Desktop\testResult\abc.txt
		//Product - Electronic - PCM Test Data - Bulk Upload
	}
}
