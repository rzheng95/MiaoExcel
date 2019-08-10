package com.rzheng.modway;

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

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class CustomsDeclarationModway {
	
	// Contract
	private final int CONTRACT_DATE_ROW = 5, CONTRACT_DATE_COL = 5;
	
	private final int QUANTITY_ROW = 15, QUANTITY_COL = 1;

	private final int TOTAL_AMOUNT_ROW = 15, TOTAL_AMOUNT_COL = 6;

	// Invoice
	private final int INVOICE_DATE_ROW = 9, INVOICE_DATE_COL = 7;

	private final int INVOICE_NUMBER_ROW = 3, INVOICE_NUMBER_COL = 5;
	
	private final int PO_ROW = 8, PO_COL = 7;

	private final int PORT_ROW = 9, PORT_COL = 6;

	
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
	private String pi_pdf_path;
	private String product_dimension_path;
	private String cd_xls_path;
	private String cd_template;
	private String invoiceNumber;
	private String invoiceDate;
	
	public CustomsDeclarationModway(String pi_pdf_path, String product_dimension_path, String cd_template, String cd_xls_path, String invoiceNumber, String invoiceDate) {
		this.error = "";
		this.pi_pdf_path = pi_pdf_path;
		this.product_dimension_path = product_dimension_path;
		this.cd_xls_path = cd_xls_path;
		this.cd_template = cd_template;
		this.invoiceNumber = invoiceNumber;
		this.invoiceDate = invoiceDate;
	}
	
	
	public String run() throws IOException {
		
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
		
		
		
		ProformaInvoiceModway pi = new ProformaInvoiceModway(pi_pdf_path);
		

		// Quantity
		double totalQuantity = pi.getTotalQuantity();
		if (totalQuantity != -1) {
			cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
			cell.setCellValue(totalQuantity);
		} else {
			error += "ERROR: Quantity not found.\n" + 
					"错误： 找不到Quantity.\n";
		}
		
		// Amount
		double totalAmount = pi.getTotalAmount();
		if (totalAmount != -1) {
			cell = worksheet.getRow(TOTAL_AMOUNT_ROW).getCell(TOTAL_AMOUNT_COL);
			cell.setCellValue(totalAmount);
		} else {
			error += "ERROR: Total Amount not found.\n" +
					"错误： 找不到Total Amount.\n";
		}
		

		
//----- Invoice page------------------------------------------------------------------
        worksheet = workbook.getSheet(Constants.INVOICE);
        
        // Invoice Date
        cell = worksheet.getRow(INVOICE_DATE_ROW).getCell(INVOICE_DATE_COL);
		cell.setCellValue(this.invoiceDate);
        
		// PO #
		String poNumber = pi.getPoNumber();
		if(poNumber != null) {
			cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
			cell.setCellValue(poNumber);
		} else {
			error += "ERROR: PO # not found.\n" +
					"错误： 找不到PO #.\n";
		}
		
		
//----- Packing List------------------------------------------------------------------	
		worksheet = workbook.getSheet(Constants.PACKING_LIST);
		
		ProductDimensionChart pdc = new ProductDimensionChart(product_dimension_path);
		
		int numOfContainers = pi.getNumberOfContainer();
		
		if (numOfContainers != -1) {
		
			double totalNetWeight = 0;
			double totalGrossWeight = 0;
			double totalCbm = 0;
			for (int i = 1; i <= numOfContainers; i++) {
				List<Item> items = pdc.getContainerItems(i);
				
				if (items != null) {
					double containerNetWeight = 0;
					double containerGrossWeight = 0;
					double cotainerCbm = 0;
					
					for (Item item : items) {
						containerNetWeight += item.getNetWeight();
						containerGrossWeight += item.getGrossWeight();
						cotainerCbm += item.getCbm();
					}
					
					totalNetWeight += containerNetWeight;
					totalGrossWeight += containerGrossWeight;
					totalCbm += cotainerCbm;
				} else {
					error += "ERROR: Items not found.\n" +
							"错误： 找不到产品.\n";
				}
			}
			
			
			// Net Weight
			cell = worksheet.getRow(NET_WEIGHT_ROW).getCell(NET_WEIGHT_COL);
			cell.setCellValue(totalNetWeight);
	
			// Gross Weight
			cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
			cell.setCellValue(totalGrossWeight);
	
			// Measurement
			cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
			cell.setCellValue(totalCbm);
		
		} else {
			error += "ERROR: Cannot get the number of containers.\n" +
					"错误： 找不到有多少个箱柜.\n";
		}

//----- 报关单------------------------------------------------------------------
		worksheet = workbook.getSheetAt(4);
		
		// Port Name (指运港)
		String port = pi.getPortName();
		if (port != null) {
			cell = worksheet.getRow(PORT_ROW).getCell(PORT_COL);
			cell.setCellValue(port);
		} else {
			error += "ERROR: Port Name Not found.\n" +
					"错误： 找不到指运港.\n";
		}
		
		
		

		if (cd_xls_path.isEmpty()) {
			cd_xls_path = invoiceNumber + "报关资料 " + poNumber + "沙发";
		} else {
			cd_xls_path += "/" + invoiceNumber + "报关资料 " + poNumber + "沙发";
		}

		
		
        
		// Close the InputStream
		if(fileInput != null)
			fileInput.close();

		if (error.isEmpty()) {
			error = "Success!";
		}
		// refreshes all formulas existed in the spreadsheet
		HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

		cd_xls_path = Util.correctFileFormat(".xls", cd_xls_path);

		// Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File(cd_xls_path));
		// write changes
		workbook.write(output_file);
		workbook.close();
		// close the stream
		output_file.close();
		
		return error;
	}
	

}





















