package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.rzheng.magnussen.ProformaInvoice;
import com.rzheng.magnussen.ShippingInstructions;
import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class CustomsDeclaration 
{
	// Contract
	private final int CONTRACT_DATE_ROW = 5, CONTRACT_DATE_COL = 5;
	
	private final int QUANTITY_ROW = 15, QUANTITY_COL = 1;

	private final int TOTAL_EXCL_TAX_ROW = 15, TOTAL_EXCL_TAX_COL = 6;

	// Invoice
	private final int INVOICE_DATE_ROW = 9, INVOICE_DATE_COL = 7;

	private final int INVOICE_NUMBER_ROW = 3, INVOICE_NUMBER_COL = 5;
	
	private final int PO_ROW = 8, PO_COL = 7;

	private final int CONSIGNEE_ROW = 7, CONSIGNEE_COL = 0;

	private final int DESTINATION_CITY_ROW = 12, DESTINATION_CITY_COL = 6;
	
	private final int DESTINATION_COUNTRY_ROW = 9, DESTINATION_COUNTRY_COL = 2;
	
	private final int RECIPIENT_ROW = 5, RECIPIENT_COL = 0;
	
	// Packing List
	private final int NET_WEIGHT_ROW = 12, NET_WEIGHT_COL = 5;
	
	private final int GROSS_WEIGHT_ROW = 12, GROSS_WEIGHT_COL = 6;
	
	private final int MEASUREMENT_ROW = 12, MEASUREMENT_COL = 7;
	
	// Read the spreadsheet that needs to be updated
	private FileInputStream fileInput;
	// Access the workbook
	private HSSFWorkbook workbook;
	// Access the worksheet, so that we can update / modify it.
	private HSSFSheet worksheet;
	// declare a Cell object
	private Cell cell;
	
	private String error;
	private String product_file_path;
	private String dimension_file_path;
	private String si_pdf_path;
	private String pi_pdf_path;
	private String cd_xls_path;
	private String cd_template;
	private String invoiceNumber;

	public CustomsDeclaration(String product_file_path, String dimension_file_path, String si_pdf_path, String pi_pdf_path, String cd_xls_path, String cd_template, String invoiceNumber) throws IOException
	{
		this.error = "";
		this.product_file_path = product_file_path;
		this.dimension_file_path = dimension_file_path;
		this.si_pdf_path = si_pdf_path;
		this.pi_pdf_path = pi_pdf_path;
		this.cd_xls_path = cd_xls_path;
		this.invoiceNumber = invoiceNumber;
		this.cd_template = cd_template;
	}
	
	public String run() throws IOException
	{
		
		try {
			this.fileInput = new FileInputStream(new File(cd_template));
		} catch (FileNotFoundException e) {
			error = "ERROR: Customs Declaration Template Not Found!\n" + this.cd_template; 
			e.printStackTrace();
			return error;
		}
		
		try {
			this.workbook = new HSSFWorkbook(fileInput);
		} catch (IOException e) {
			error = "ERROR: FileInputStream Exception";
			e.printStackTrace();
			return error;
		}
		
		
		
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date current_date = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(current_date);
        calendar.add(Calendar.MONTH, -2);
        
        Date two_months_ago = calendar.getTime();
        
        String invoice_date = dateFormat.format(current_date);
        String contract_date = dateFormat.format(two_months_ago);
        

        if (invoiceNumber.trim().isEmpty()) {
        	invoiceNumber = "INYB" + calendar.get(Calendar.YEAR) + "US" + calendar.get(Calendar.MONTH) + calendar.get(Calendar.DAY_OF_MONTH);
        }
        
//----- Contract page------------------------------------------------------------------
        worksheet = workbook.getSheet(Constants.CONTRACT);
        
        // Invoice number
        cell = worksheet.getRow(INVOICE_NUMBER_ROW).getCell(INVOICE_NUMBER_COL);
		cell.setCellValue(invoiceNumber);
		
		// Contract Date
        cell = worksheet.getRow(CONTRACT_DATE_ROW).getCell(CONTRACT_DATE_COL);
		cell.setCellValue(contract_date);
		
		
		ProformaInvoice pi = new ProformaInvoice(product_file_path, dimension_file_path, pi_pdf_path);
		

		// Quantity
		double quantity = pi.getQuantity();
		if (quantity != -1) {
			cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
			cell.setCellValue(quantity);
		} else {
			error += "ERROR: Quantity not found.\n" + 
					"错误： 找不到Quantity.\n";
		}
		
		// Amount
		double totalExclTaxAmount = pi.getTotalExclTaxAmount();
		if (totalExclTaxAmount != -1) {
			cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
			cell.setCellValue(totalExclTaxAmount);
		} else {
			error += "ERROR: Total Excl. Tax not found.\n" +
					"错误： 找不到Total Excl. Tax.\n";
		}
		
		
		
        ShippingInstructions si = new ShippingInstructions(si_pdf_path);
        
        
//----- Invoice page------------------------------------------------------------------
        worksheet = workbook.getSheet(Constants.INVOICE);
        
        // Invoice Date
        cell = worksheet.getRow(INVOICE_DATE_ROW).getCell(INVOICE_DATE_COL);
		cell.setCellValue(invoice_date);
		
		// Consignee	
		String consignee = si.getConsignee();
		if(consignee != null) {
			cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
			cell.setCellValue(consignee);
		} else {
			error += "ERROR: Consignee not found.\n" +
					"错误： 找不到Consignee.\n";
		}
		
		// PO #
		String poNumber = si.getPoNumber();
		if(poNumber != null) {
			cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
			cell.setCellValue(poNumber.substring(Constants.PO.length()).trim());
		} else {
			error += "ERROR: PO # not found.\n" +
					"错误： 找不到PO #.\n";
		}
        
		// Destination City
		String city = si.getDestinationCity();
		if(city != null) {
			cell = worksheet.getRow(DESTINATION_CITY_ROW).getCell(DESTINATION_CITY_COL);
   		 	cell.setCellValue(city);
		} else {
			error += "ERROR: Destination City not found.\n" +
					"错误： 找不到Destination City.\n";
		}
		
		
//----- Packing List------------------------------------------------------------------	
		worksheet = workbook.getSheet(Constants.PACKING_LIST);
		List<Object> stats = pi.getTotalStats(pi.getItems());
		
		if (stats != null) {
			if (stats.get(Constants.ERROR_CODE_INDEX).toString().isEmpty()) {
				// Net Weight
				cell = worksheet.getRow(NET_WEIGHT_ROW).getCell(NET_WEIGHT_COL);
				cell.setCellValue(Double.parseDouble(stats.get(Constants.TOTAL_NET_WEIGHT_INDEX).toString()));

				// Gross Weight
				cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
				cell.setCellValue(Double.parseDouble(stats.get(Constants.TOTAL_GROSS_WEIGHT_INDEX).toString()));

				// Measurement
				cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
				cell.setCellValue(Double.parseDouble(stats.get(Constants.TOTAL_VOLUME_INDEX).toString()));
			} else {
				error = stats.get(Constants.ERROR_CODE_INDEX).toString();
			}
		}

//----- 报关单------------------------------------------------------------------
		worksheet = workbook.getSheet(Constants.BAO_GUAN_DAN);
		// Recipient (First line of Consignee)
		if (consignee != null) {
			cell = worksheet.getRow(RECIPIENT_ROW).getCell(RECIPIENT_COL);
			cell.setCellValue(consignee.split("\\r?\\n")[0]);
		}
		
		// Country 贸易国/运抵国
		String country = si.getDestinationCountry();
		if (country != null) {
			if (country.equalsIgnoreCase(Constants.SAUDI_ARABIA))
				country = "沙特阿拉伯";
			cell = worksheet.getRow(DESTINATION_COUNTRY_ROW).getCell(DESTINATION_COUNTRY_COL);
			cell.setCellValue(country);
		} else {
			error += "ERROR: Destination Country not found.\n" +
					"错误： 找不到Destination Country.\n";
		}
        
        
		if (cd_xls_path.trim().isEmpty()) {
			String[] poNum = si.getPoNumber().split(" ");
			if (poNum != null && poNum.length >= 3)
				cd_xls_path = invoiceNumber + " SOFA " + poNum[2] ; // + PO
        }
		
		
		// Close the InputStream
		if(fileInput != null)
			fileInput.close();

		if (error.isEmpty()) {
			error = "Success!";

			// refreshes all formulas existed in the spreadsheet
			HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

			cd_xls_path = Util.correctXlsFilename(cd_xls_path);

			// Open FileOutputStream to write updates
			FileOutputStream output_file = new FileOutputStream(new File(cd_xls_path));
			// write changes
			workbook.write(output_file);
			// close the stream
			output_file.close();
		}

		
		
		
		return this.error;
	}
	
	

}































