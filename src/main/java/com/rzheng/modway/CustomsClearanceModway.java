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
	private final int HOUSE_BILL_NUMBER_ROW = 9, HOUSE_BILL_NUMBER_COL = 10;
	private final int CONTAINER_NUMBERS_ROW = 11, CONTAINER_NUMBERS_COL = 9;
	private final int PO_NUMBER_ROW = 16, PO_NUMBER_COL = 1;
	
	private final int TO_ROW = 8, TO_COL = 2;
	
	private final int ETD_ROW = 16, ETD_COL = 9;
	private final int ETA_ROW = 16, ETA_COL = 10;
	
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
	private String etd;
	private String eta;
	
	public CustomsClearanceModway(String pi_path, String oceanBillOfLading_path, String product_dimension_path, String cc_template, String cc_xlsx_path, String invoiceNumber,
			String etd, String eta) {
		this.error = "";
		this.pi_path = pi_path;
		this.oceanBillOfLading_path = oceanBillOfLading_path;
		this.product_dimension_path = product_dimension_path;
		this.cc_template = cc_template;
		this.cc_xlsx_path = cc_xlsx_path;
		this.invoiceNumber = invoiceNumber;
		this.etd = etd;
		this.eta = eta;
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
        
        // Master CI
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
			error += "ERROR: Container No. not found in the given PI.\n" + 
					"错误： 找不到Container No.\n";
		}
		
		OceanBillOfLadingModway oblm = new OceanBillOfLadingModway(oceanBillOfLading_path);
		
		// Container #
		String containerNumbers = oblm.getAllContainerNumbers();
		if (containerNumbers != null) {
			cell = worksheet.getRow(CONTAINER_NUMBERS_ROW).getCell(CONTAINER_NUMBERS_COL);
			cell.setCellValue(containerNumbers);
		} else {
			error += "ERROR: Container Numbers not found in the given Ocean Bill of Lading.\n" + 
					"错误： 找不到Container Numbers.\n";
		}
		
		// House Bill # (Bill of Lading No.)
		String houseBillNumber = oblm.getBillOfLadingNumber();
		if (houseBillNumber != null) {
			cell = worksheet.getRow(HOUSE_BILL_NUMBER_ROW).getCell(HOUSE_BILL_NUMBER_COL);
			cell.setCellValue(houseBillNumber);
		} else {
			error += "ERROR: Bill of Lading No. not found.\n" + 
					"错误： 找不到Bill of Lading No.\n";
		}
		
		
		// PO #
		String poNumber = pi.getPoNumber();
		if (poNumber != null) {
			cell = worksheet.getRow(PO_NUMBER_ROW).getCell(PO_NUMBER_COL);
			cell.setCellValue(poNumber);
			
			workbook.setSheetName(0, poNumber + " MASTER CI");
			if (this.cc_xlsx_path.isEmpty()) {
				this.cc_xlsx_path = poNumber + " CI & PL";
			}
		} else {
			error += "ERROR: Purcharse Order No. (PO #) not found.\n" + 
					"错误： 找不到Purcharse Order No. (PO #).\n";
		}
		
		// To (Place of Discharge)
		String to = oblm.getPlaceOfDischarge();
		if (to != null) {
			cell = worksheet.getRow(TO_ROW).getCell(TO_COL);
			cell.setCellValue(to);
		} else {
			error += "ERROR: Place of Discharge not found.\n" + 
					"错误： 找不到Place of Discharge.\n";
		}
		
		// ETD
		if (!this.etd.isEmpty()) {
			cell = worksheet.getRow(ETD_ROW).getCell(ETD_COL);
			cell.setCellValue(this.etd);
		}
		
		// ETA
		if (!this.eta.isEmpty()) {
			cell = worksheet.getRow(ETA_ROW).getCell(ETA_COL);
			cell.setCellValue(this.eta);
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



















