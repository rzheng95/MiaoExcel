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

import com.rzheng.excel.util.Constants;
import com.rzheng.excel.util.Util;

public class CustomsClearance 
{
	// Factory Invoice - RDC & ADC
	private final int SHIP_TO_NAME_ROW = 15, SHIP_TO_NAME_COL = 1;
	private final int SHIP_TO_ADDRESS_ROW = 16, SHIP_TO_ADDRESS_COL = 1;
	private final int FORWARDER_NAME_ROW = 21, FORWARDER_NAME_COL = 1;
	private final int CARRIER_ROW = 22, CARRIER_COL = 1;
	
	
	private final int DATE_ROW = 9, DATE_COL = 7;
	private final int INVOICE_ROW = 9, INVOICE_COL = 7;
	
	// Contract
	private final int CONTRACT_DATE_ROW = 5, CONTRACT_DATE_COL = 5;
	
	private final int QUANTITY_ROW = 15, QUANTITY_COL = 1;

	private final int TOTAL_EXCL_TAX_ROW = 15, TOTAL_EXCL_TAX_COL = 6;

	// Invoice
	

	private final int INVOICE_NUMBER_ROW = 5, INVOICE_NUMBER_COL = 6;
	
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
	private String cc_xls_path;
	private String cc_template;
	private String invoiceNumber;

	public CustomsClearance(String product_file_path, String dimension_file_path, String si_pdf_path, String pi_pdf_path, String cc_xls_path, String cc_template, String invoiceNumber) throws IOException
	{
		this.error = "";
		this.product_file_path = product_file_path;
		this.dimension_file_path = dimension_file_path;
		this.si_pdf_path = si_pdf_path;
		this.pi_pdf_path = pi_pdf_path;
		this.cc_xls_path = cc_xls_path;
		this.cc_template = cc_template;
		this.invoiceNumber = invoiceNumber;
		

	}
	
	public String run() throws IOException
	{
		
		try {
			this.fileInput = new FileInputStream(new File(cc_template));
		} catch (FileNotFoundException e) {
			error = "ERROR: Customs Declaration Template Not Found!\n" + this.cc_template; 
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


        String invoice_date = dateFormat.format(current_date);

        

        if (invoiceNumber.trim().isEmpty()) {
        	invoiceNumber = "INYB" + calendar.get(Calendar.YEAR) + "US" + calendar.get(Calendar.MONTH) + calendar.get(Calendar.DAY_OF_MONTH);
        }
        
        
        
//----- Factory Invoice - RDC & ADC page------------------------------------------------------------------
        worksheet = workbook.getSheetAt(0);
        


        // Invoice number
        cell = worksheet.getRow(INVOICE_NUMBER_ROW).getCell(INVOICE_NUMBER_COL);
		cell.setCellValue(invoiceNumber);
		
		ShippingInstructions si = new ShippingInstructions(si_pdf_path);
		
		ProformaInvoice pi = new ProformaInvoice(product_file_path, dimension_file_path, pi_pdf_path);
		
		
        // Ship-to Name (Consignee)
        String consignee = si.getConsignee();
		if (consignee != null) {
			cell = worksheet.getRow(SHIP_TO_NAME_ROW).getCell(SHIP_TO_NAME_COL);
			cell.setCellValue(consignee);
		} else {
			error = "ERROR: Consignee not found.\n" + 
					"错误： 找不到Consignee.\n";
		}
		
		// Ship-to Address
		String shipToAddress = si.getShipToAddress();
		if (shipToAddress != null) {
			cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
			cell.setCellValue(shipToAddress);
		} else {
			error = "ERROR: Ship-to Address not found.\n" + 
					"错误： 找不到Ship-to Address.\n";
		}
		
		// Freight Forwarder Name (first line)
		String forwarderName = si.getForwarder();
		if (forwarderName != null) {
			String[] lines = forwarderName.split("\\r?\\n");
			if(lines != null && lines.length >= 1) {
				cell = worksheet.getRow(FORWARDER_NAME_ROW).getCell(FORWARDER_NAME_COL);
				cell.setCellValue(lines[0]);
			}	
		} else {
			error = "ERROR: Forwarder Name not found.\n" + 
					"错误： 找不到Forwarder Name.\n";
		}
		
		// Carrier Name
		String carrier = si.getCarrier();
		if (carrier != null) {
			cell = worksheet.getRow(CARRIER_ROW).getCell(CARRIER_COL);
			cell.setCellValue(carrier);

		} else {
			error = "ERROR: Carrier not found.\n" + 
					"错误： 找不到Carrier.\n";
		}
		
		

        
//----- Invoice page------------------------------------------------------------------
        worksheet = workbook.getSheet(Constants.INVOICE);
        
        // Invoice Date
        //cell = worksheet.getRow(INVOICE_DATE_ROW).getCell(INVOICE_DATE_COL);
		cell.setCellValue(invoice_date);
		
		
		// PO #
		cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
		String poNumber = si.getPoNumber();
		if(poNumber != null) {
			cell.setCellValue(poNumber.substring(Constants.PO.length()).trim());
		} else {
			error = "ERROR: PO # not found.\n" +
					"错误： 找不到PO #.\n";
		}
        
		// Destination City
		cell = worksheet.getRow(DESTINATION_CITY_ROW).getCell(DESTINATION_CITY_COL);
		String city = si.getDestinationCity();
		if(city != null) {
   		 	cell.setCellValue(city);
		} else {
			error = "ERROR: Destination City not found.\n" +
					"错误： 找不到Destination City.\n";
		}
		

        
        
		if (cc_xls_path.trim().isEmpty()) {
			String[] poNum = si.getPoNumber().split(" ");
			if (poNum != null && poNum.length >= 3)
				cc_xls_path = invoiceNumber + " SOFA " + poNum[2] ; // + PO
        }
		
		
		// Close the InputStream
		if(fileInput != null)
			fileInput.close();

		if (error.isEmpty()) {
			error = "Success!";

			// refreshes all formulas existed in the spreadsheet
			HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

			cc_xls_path = Util.correctXlsFilename(cc_xls_path);

			// Open FileOutputStream to write updates
			FileOutputStream output_file = new FileOutputStream(new File(cc_xls_path));
			// write changes
			workbook.write(output_file);
			// close the stream
			output_file.close();
		}

		
		
		
		return this.error;
	}
	
	

}































