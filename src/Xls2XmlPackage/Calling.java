package Xls2XmlPackage;

import java.util.Scanner;

public class Calling {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Xls2XmlApparel obj=new Xls2XmlApparel();
//		System.out.println(args[0]+" "+args[1]);
		Scanner in = new Scanner(System.in);
//		in.next();
		while(true){
		System.out.println("Enter the input file path with name ex: C:\\excel2xmldata\\input\\excel2xml.xls");
		String input=in.nextLine().trim();
		System.out.println("Enter the output folder path ex: C:\\excel2xmldata\\output");
		String output=in.nextLine().trim();		
		obj.mainf(input,output);
		
		System.out.println("Do you want to process a new file? press (Y/N)");
		char more = in.next().charAt(0);
		if(more =='y' || more =='Y'){
			
		}else{
			break;
		}
		}
	}

}
