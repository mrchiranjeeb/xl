package Xls2XmlPackage;

import java.io.File;
import java.io.FileOutputStream;

import jxl.Sheet;
import jxl.Workbook;

public class CsvConverter {
	public void Apparelcsv(String inFileName, String outFolderName)
			throws Exception {
		File f = new File(inFileName);
		Workbook wb = null;
		try {
			wb = Workbook.getWorkbook(f);
			// System.out.println("workbook loaded");
		} catch (Exception e) {
			System.out
					.println("Input file or path is invalid.. please check..");
			return;
		}
		Sheet s = wb.getSheet(0);
		int rowNum = s.getRows();
		int colNum = s.getColumns();
		int iteration = 6, inneriter;
		StringBuffer celldata = new StringBuffer(
				"PRODUCTCODE,PRODUCTFEATUREQUALIFIER,PRODUCTFEATUREVALUE,");
		String temp;
		for (; iteration < rowNum; iteration++) {
			// temp=s.getCell(4,3).getContents().trim();
			// System.out.println(temp);
			celldata.append("\r\n");
			for (inneriter = 12; inneriter <= 126; inneriter++) {
				if ((s.getCell(inneriter, iteration).getContents().trim())
						.length() > 0) {
					if (inneriter == 126) {
						temp = s.getCell(4, iteration).getContents().trim();
						System.out.println(temp);
						celldata.append(temp
								+ ","
								+ s.getCell(inneriter, 3).getContents().trim()
								+ ","
								+ s.getCell(inneriter, iteration).getContents()
										.trim().replaceAll(",", " "));
					} else {
						temp = s.getCell(4, iteration).getContents().trim();

						celldata.append(temp
								+ ","
								+ s.getCell(inneriter, 3).getContents().trim()
								+ ","
								+ s.getCell(inneriter, iteration).getContents()
										.trim().replaceAll(",", " ") + "\r\n");
					}
					// System.out.println(celldata);
				}
			}

		}
		FileOutputStream fos = new FileOutputStream(outFolderName);
		fos.write(celldata.toString().getBytes());
		fos.close();

	}

	public void ElectronicCsv(String inFileName, String outFolde)
			throws Exception {
		File f = new File(inFileName);
		Workbook wb = null;
		try {
			wb = Workbook.getWorkbook(f);
			// System.out.println("workbook loaded");
		} catch (Exception e) {
			System.out
					.println("Input file or path is invalid.. please check..");
			return;
		}
		Sheet s = wb.getSheet(0);
		int rowNum = s.getRows();
		int colNum = s.getColumns();
		int iteration = 6, inneriter;
		StringBuffer celldata = new StringBuffer(
				"PRODUCTCODE,PRODUCTFEATUREQUALIFIER,PRODUCTFEATUREVALUE,");
		String temp, temp2;
		System.out.println(rowNum + " " + colNum);
		for (; iteration < rowNum; iteration++) {
			try {
				// temp=s.getCell(4,3).getContents().trim();
				// System.out.println(temp);
				celldata.append("\r\n");
				for (inneriter = 12; inneriter <= 255; inneriter++) {
					if ((s.getCell(inneriter, iteration).getContents().trim())
							.length() > 0) {
						if (inneriter == 255) {
							temp = s.getCell(4, iteration).getContents().trim();
							// System.out.println(temp);
							System.out.println(temp
									+ " "
									+ s.getCell(inneriter, 3).getContents()
											.trim());
							celldata.append(temp
									+ ","
									+ s.getCell(inneriter, 3).getContents()
											.trim()
									+ ","
									+ s.getCell(inneriter, iteration)
											.getContents().trim()
											.replaceAll("[,\\n\\r]+", " "));
						} else {
							temp = s.getCell(4, iteration).getContents().trim();
							System.out.println(temp
									+ " "
									+ s.getCell(inneriter, 3).getContents()
											.trim());
							celldata.append(temp
									+ ","
									+ s.getCell(inneriter, 3).getContents()
											.trim()
									+ ","
									+ s.getCell(inneriter, iteration)
											.getContents().trim()
											.replaceAll("[,\\n\\r]+", " ")
									+ "\r\n");
						}
						// System.out.println(celldata);
					}

				}
				// ====================================================================================
				// ========================================Sheet2
				// starts===============================
				// ====================================================================================
				s = wb.getSheet(1);
				for (inneriter = 4; inneriter <= 39; inneriter++) {
					if ((s.getCell(inneriter, iteration).getContents().trim())
							.length() > 0) {
						if (inneriter == 39) {
							if ((s.getCell(inneriter, iteration).getContents()
									.trim()).length() <= 0) {

							} else {
								temp2 = wb.getSheet(0).getCell(4, iteration)
										.getContents().trim();
								System.out.println(temp2
										+ " "
										+ s.getCell(inneriter, 3).getContents()
												.trim());
								celldata.append(temp2
										+ ","
										+ s.getCell(inneriter, 3).getContents()
												.trim()
										+ ","
										+ s.getCell(inneriter, iteration)
												.getContents().trim()
												.replaceAll("[,\\n\\r]+", " "));
							}
						} else {
							temp2 = wb.getSheet(0).getCell(4, iteration)
									.getContents().trim();
							System.out.println(temp2
									+ " "
									+ s.getCell(inneriter, 3).getContents()
											.trim());
							celldata.append(temp2
									+ ","
									+ s.getCell(inneriter, 3).getContents()
											.trim()
									+ ","
									+ s.getCell(inneriter, iteration)
											.getContents().trim()
											.replaceAll("[,\\n\\r]+", " ")
									+ "\r\n");
						}
						// System.out.println(celldata);
					}
				}
				s = wb.getSheet(0);
				inneriter = 0;
			} catch (Exception e) {

			}
		}
		FileOutputStream fok = new FileOutputStream(outFolde);
		fok.write(celldata.toString().getBytes());
		fok.close();

	}

}
