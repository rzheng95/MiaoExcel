package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Scanner;

import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import io.github.jonathanlink.PDFLayoutTextStripper;

public class CustomsDeclaration 
{
	private final String TOTAL = "TOTAL";
	private final int QUANTITY_ROW = 15, QUANTITY_COL = 1;
	
	private final String SUB_TOTAL = "SUB TOTAL";
	
	private final String TOTAL_EXCL_TAX = "TOTAL EXCL. TAX";
	private final int TOTAL_EXCL_TAX_ROW = 15, TOTAL_EXCL_TAX_COL = 6;
	
	private final String PO = "P.O.NO.";
	private final int PO_SHEET = 1, PO_ROW = 8, PO_COL = 7;
	
	private final String CONSIGNEE = "CONSIGNEE:";
	private int CONSIGNEE_SHEET = 1, CONSIGNEE_ROW = 7, CONSIGNEE_COL = 0;
	private final String NOTIFY = "NOTIFY: ";
	
	private final String DESTINATION = "DESTINATION:";
	private final int DESTINATION_SHEET = 1, DESTINATION_ROW = 12, DESTINATION_COL = 6;
	
	// Read the spreadsheet that needs to be updated
	FileInputStream fileInput = new FileInputStream(new File("Customs Declaration Template.xls"));
	// Access the workbook
	HSSFWorkbook workbook = new HSSFWorkbook(fileInput);
	// Access the worksheet, so that we can update / modify it.
	HSSFSheet worksheet;

	// declare a Cell object
	Cell cell = null;

	public CustomsDeclaration(String si_pdf_path, String pi_pdf_path, String cd_xls_path) throws IOException
	{
		if(!si_pdf_path.isEmpty() && !pi_pdf_path.isEmpty())
			this.contract(si_pdf_path, pi_pdf_path, cd_xls_path);
	}
	
	public void contract(String si_pdf_path, String pi_pdf_path, String cd_xls_path) throws IOException
	{
		try (PDDocument pi = PDDocument.load(new File(pi_pdf_path));
				PDDocument si = PDDocument.load(new File(si_pdf_path))) {
			pi.getClass();
			si.getClass();
			
			if (!pi.isEncrypted() && !si.isEncrypted()) {

				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);

				PDFTextStripper tStripper = new PDFTextStripper();

				String piText = tStripper.getText(pi);
				String siText = tStripper.getText(si);
				
				
				// SI
				String[] lines = siText.split("\\r?\\n");
				for (int i = 0; i < lines.length; i++)
					lines[i] = lines[i].toUpperCase();
				int i = 0;
				
				while (i < lines.length) {

					worksheet = workbook.getSheetAt(0);
					if(lines[i].contains(CONSIGNEE)) {

						worksheet = workbook.getSheetAt(CONSIGNEE_SHEET);
						// Access the second cell in second row to update the value
	            		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
	            		// Get current cell value value and overwrite the value
	           		 	cell.setCellValue(lines[i].substring(CONSIGNEE.length()).trim());
	           		 	i++;
                		while (!lines[i].contains(NOTIFY))
                		{
                			if (lines[i].trim().isEmpty())
                			{
                				i++;
                				continue;
                			}
                			CONSIGNEE_ROW++;
	                		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
		           		 	cell.setCellValue(lines[i].trim());

		           		 	i++;
                		}
                		i--;
					}
					else if (lines[i].contains(DESTINATION)) {
						worksheet = workbook.getSheetAt(DESTINATION_SHEET);
						String city = lines[i].substring(DESTINATION.length()).trim();
						if(city.contains(",")) {
							String[] arr = city.split(",");
							city = arr[0].trim();	
						}
						
						cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
	           		 	cell.setCellValue(city);
					}
					
					
					
					i++;
				}
				
				
				
				// PI
				lines = piText.split("\\r?\\n");
				
				i = 0;

				while (i < lines.length) {
					
					worksheet = workbook.getSheetAt(0);
					
					if (lines[i].contains(TOTAL)) {
						if (lines[i].contains(TOTAL_EXCL_TAX)) {
		            		cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
		            		double amount = extractNumberFromString(lines[i].substring(TOTAL_EXCL_TAX.length()));
		           		 	cell.setCellValue(amount);
//		            		cell.setCellType(CellType.NUMERIC);
//		            		CellStyle cs = wb.createCellStyle();
//		            		cs.setDataFormat((short)7);
//		            		cell.setCellStyle(cs);
							
		           		 	// set total amount
		           		 	cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
		           		 	cell.setCellValue(amount);
		           		 	i++;
		           		 	continue;
		           		 	
						} else if (lines[i].contains(SUB_TOTAL)) {
							i++;
							continue;
						}
						String[] arr = lines[i].split(" ");
						cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
						cell.setCellValue(Integer.parseInt(arr[arr.length-1].trim()));

					}
					else if (lines[i].contains(PO))
					{
						String po_num = lines[i].substring(lines[i].indexOf(PO) + PO.length()).trim();
						worksheet = workbook.getSheetAt(PO_SHEET);
						cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
						cell.setCellValue(po_num);
					}

					i++;
				}

			}

		}
		HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
		// Close the InputStream
		fileInput.close();

		if (!cd_xls_path.contains(".xls") && !cd_xls_path.isEmpty())
			cd_xls_path = cd_xls_path + ".xls";
		if (cd_xls_path.contains(".xlsx"))
			cd_xls_path = cd_xls_path.substring(0, cd_xls_path.length() - 1);

		// Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File(cd_xls_path));
		// write changes
		workbook.write(output_file);
		// close the stream
		output_file.close();
	}
	
	private double extractNumberFromString(String amount) {
		amount = amount.replaceAll("\\$", "").trim();
		amount = amount.replaceAll(",", "").trim();
		return Double.parseDouble(amount);
	}
	
	private boolean countDollarSign(String line, int countThreshold) {
		int count = 0;
		StringBuilder sb = new StringBuilder(line);
		for (int i = 0; i < countThreshold; i++) {
			if(sb.indexOf("$") != -1) {
				sb = sb.deleteCharAt(line.indexOf('$'));
				count++;
			}
			if(count >= countThreshold)
				return true;
			
		}
		return false;	
	}
}































