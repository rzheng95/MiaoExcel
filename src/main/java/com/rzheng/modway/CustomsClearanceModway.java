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

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class CustomsClearanceModway 
{
	private final int DATE_ROW = 2, DATE_COL = 10;
	private final int INVOICE_NUMBER_ROW = 4, INVOICE_NUMBER_COL = 10;
	private final int CONTAINER_QTY_ROW = 8, CONTAINER_QTY_COL = 10;
	
	
	// Read the spreadsheet that needs to be updated
	private FileInputStream fileInput;
	// Access the workbook
	private XSSFWorkbook workbook;
	// Access the worksheet, so that we can update / modify it.
	private XSSFSheet worksheet;
	// declare a Cell object
	private Cell cell;
	
	
	private String error;
	private String pi_path;
	private String oceanBillOfLading_path;
	private String product_dimension_path;
	private String cc_template;
	private String cc_xlsx_path;
	private String invoiceNumber;
	
	public CustomsClearanceModway(String pi_path, String oceanBillOfLading_path, String product_dimension_path, String cc_template, String cc_xlsx_path, String invoiceNumber) {
		this.error = "";
		this.pi_path = pi_path;
		this.oceanBillOfLading_path = oceanBillOfLading_path;
		this.product_dimension_path = product_dimension_path;
		this.cc_template = cc_template;
		this.cc_xlsx_path = cc_xlsx_path;
		this.invoiceNumber = invoiceNumber;
	}
	
	
	public String run() throws IOException {
		
		try {
			this.fileInput = new FileInputStream(new File(cc_template));
		} catch (FileNotFoundException e) {
			error = "ERROR: Customs Declaration Template Not Found!\n" + this.cc_template; 
			e.printStackTrace();
			return error;
		}
		
		try {
			this.workbook = new XSSFWorkbook(fileInput);
		} catch (IOException e) {
			error = "ERROR: FileInputStream Exception";
			e.printStackTrace();
			return error;
		}
		
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date current_date = new Date();
        Calendar calendar = Calendar.getInstance();
        
        worksheet = workbook.getSheetAt(0);
        
        // Date
        cell = worksheet.getRow(DATE_ROW).getCell(DATE_COL);
		cell.setCellValue(dateFormat.format(current_date));
		

		// Invoice #
		if (this.invoiceNumber == null || this.invoiceNumber.isEmpty()) {
			this.invoiceNumber = "INYB" + calendar.get(Calendar.YEAR) + "US" + calendar.get(Calendar.MONTH) + calendar.get(Calendar.DAY_OF_MONTH);
		}
		
		cell = worksheet.getRow(INVOICE_NUMBER_ROW).getCell(INVOICE_NUMBER_COL);
		cell.setCellValue(this.invoiceNumber);
		
		ProformaInvoiceModway pi = new ProformaInvoiceModway(pi_path);
		
		// Container Qty
		String containerQty = pi.getContainerQty();
		if (containerQty != null) {
			cell = worksheet.getRow(CONTAINER_QTY_ROW).getCell(CONTAINER_QTY_COL);
			cell.setCellValue(containerQty);
		} else {
			error += "ERROR: Container No. not found.\n" + 
					"错误： 找不到Container No.\n";
		}
		
		
		
		
		
		
		
		if (error.isEmpty()) {
			error = "Success!";

			// refreshes all formulas existed in the spreadsheet
			XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

			cc_xlsx_path = Util.correctXlsxFilename(cc_xlsx_path);

			// Open FileOutputStream to write updates
			FileOutputStream output_file;
			
			try {
				output_file = new FileOutputStream(new File(cc_xlsx_path));
				// write changes
				workbook.write(output_file);
				// close the stream
				output_file.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		
		return error;
	}
}



















