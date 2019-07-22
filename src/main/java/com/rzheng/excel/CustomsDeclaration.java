package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class CustomsDeclaration 
{
	
	private final int CONTRACT_DATE_ROW = 5, CONTRACT_DATE_COL = 5;

	private final int INVOICE_DATE_ROW = 9, INVOICE_DATE_COL = 7;

	private final int INVOICE_NUMBER_ROW = 3, INVOICE_NUMBER_COL = 5;

	private final int QUANTITY_ROW = 15, QUANTITY_COL = 1;

	private final int TOTAL_EXCL_TAX_ROW = 15, TOTAL_EXCL_TAX_COL = 6;

	private final int PO_ROW = 8, PO_COL = 7;

	private final int CONSIGNEE_ROW = 7, CONSIGNEE_COL = 0;

	private final int DESTINATION_ROW = 12, DESTINATION_COL = 6;
	
	private final int RECIPIENT_SHEET = 4, RECIPIENT_ROW = 5, RECIPIENT_COL = 0;
	
	// Read the spreadsheet that needs to be updated
	FileInputStream fileInput = new FileInputStream(new File("Customs Declaration Template.xls"));
	// Access the workbook
	HSSFWorkbook workbook = new HSSFWorkbook(fileInput);
	// Access the worksheet, so that we can update / modify it.
	HSSFSheet worksheet;

	// declare a Cell object
	Cell cell = null;
	
	

	public CustomsDeclaration(String si_pdf_path, String pi_pdf_path, String cd_xls_path, String invoiceNumber) throws IOException
	{
		if(!si_pdf_path.isEmpty() && !pi_pdf_path.isEmpty()) {
			this.contract(si_pdf_path, pi_pdf_path, cd_xls_path, invoiceNumber);
		}
			
	}
	
	public void contract(String si_pdf_path, String pi_pdf_path, String cd_xls_path, String invoiceNumber) throws IOException
	{
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date current_date = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(current_date);
        calendar.add(Calendar.MONTH, -2);
        
        Date two_months_ago = calendar.getTime();
        
        String invoice_date = dateFormat.format(current_date);
        String contract_date = dateFormat.format(two_months_ago);
        

        if (invoiceNumber.trim().isEmpty()) {
        	invoiceNumber = "INYB" + invoice_date;
        }
        worksheet = workbook.getSheet(Constants.CONTRACT);
        cell = worksheet.getRow(INVOICE_NUMBER_ROW).getCell(INVOICE_NUMBER_COL);
		cell.setCellValue(invoiceNumber);
		
        cell = worksheet.getRow(CONTRACT_DATE_ROW).getCell(CONTRACT_DATE_COL);
		cell.setCellValue(contract_date);
		
		worksheet = workbook.getSheet(Constants.INVOICE);
        cell = worksheet.getRow(INVOICE_DATE_ROW).getCell(INVOICE_DATE_COL);
		cell.setCellValue(invoice_date);
		
        
        
        
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
//				System.out.println(siText);
				// SI
				String[] lines = siText.split("\\r?\\n");
				int i = 0;
				
				while ( i < lines.length ) {

					worksheet = workbook.getSheetAt(0);
					
					if(lines[i].toUpperCase().contains(Constants.CONSIGNEE)) {
						
						String str = lines[i].substring(Constants.CONSIGNEE.length()).trim();
						
						worksheet = workbook.getSheetAt(RECIPIENT_SHEET);
                		cell = worksheet.getRow(RECIPIENT_ROW).getCell(RECIPIENT_COL);
                		cell.setCellValue(str);
						
	            		
                		worksheet = workbook.getSheet(Constants.INVOICE);
	            		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);

	           		 	
	           		 	i++;
                		while (!Util.checkEndLine(lines[i].toUpperCase()))
                		{
                			if (lines[i].trim().isEmpty())
                			{
                				i++;
                				continue;
                			}
                			str += "\n" + lines[i].trim();
		           		 	i++;
                		}
                		i--;
                		cell.setCellValue(str);
                		
                		
                		
					}
					else if (lines[i].toUpperCase().contains(Constants.DESTINATION)) {

						worksheet = workbook.getSheet(Constants.INVOICE);
						String city = lines[i].substring(Constants.DESTINATION.length()).trim();

						if (city.contains(",")) {
							String[] arr = city.split(",");
							city = arr[0].trim();	
						} else if  (city.contains("-")) {
							String[] arr = city.split("-");
							city = arr[0].trim();	
						} else if  (city.contains("–")) {
							String[] arr = city.split("–");
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
					
					if (lines[i].contains(Constants.TOTAL)) {
						if (lines[i].contains(Constants.TOTAL_EXCL_TAX)) {
		            		cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
		            		double amount = extractNumberFromString(lines[i].substring(Constants.TOTAL_EXCL_TAX.length()));
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
		           		 	
						} else if (lines[i].contains(Constants.SUB_TOTAL)) {
							i++;
							continue;
						}
						String[] arr = lines[i].split(" ");
						cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
						cell.setCellValue(Integer.parseInt(arr[arr.length-1].trim()));

					}
					else if (lines[i].contains(Constants.PI_PO))
					{
						String po_num = lines[i].substring(lines[i].indexOf(Constants.PI_PO) + Constants.PI_PO.length()).trim();
						worksheet = workbook.getSheet(Constants.INVOICE);
						cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
						cell.setCellValue(po_num);
					}

					i++;
				}

			}
			si.close();
			pi.close();
		}
		
		if (cd_xls_path.trim().isEmpty()) {
        	cd_xls_path = invoiceNumber + " " ; // + PO
        }
		
		// refreshes all formulas existed in the spreadsheet
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

}































