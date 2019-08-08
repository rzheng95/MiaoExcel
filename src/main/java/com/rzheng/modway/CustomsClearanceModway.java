package com.rzheng.modway;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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
	
	private int item_row = 20;
	private final int STYLE_NUMBER_COL = 1;
	private final int VENDOR_STYLE_NUMBER_COL = 2;
	private final int DESCRIPTION_COL = 3;
	private final int MATERIAL_COL = 4;
	private final int HTS_CODE_COL = 5;
	private final int QUANTITY_COL = 6;
	private final int CARTON_COL = 7;
	private final int PACKAGE_PRICE_COL = 8;
	private final int UNIT_PRICE_COL = 9;
	private final int TOTAL_AMOUNT_COL = 10;
	
	private final int NET_WEIGHT_COL = 8;
	private final int GROSS_WEIGHT_COL = 9;
	private final int CMB_COL = 10;
	
	// Read the spreadsheet that needs to be updated
	private FileInputStream fileInput;
	// Access the workbook
	private Workbook workbook;
	// Access the worksheet, so that we can update / modify it.
	private Sheet worksheet;
	// declare a Cell object
	private Cell cell;
	
	
	private String error;
	private String pi_path;
	private String product_dimension_chart_path;
	private String oceanBillOfLading_path;
	private String cc_template;
	private String cc_xls_path;
	private String invoiceNumber;
	private String etd;
	private String eta;
	
	public CustomsClearanceModway(String pi_path, String oceanBillOfLading_path, String product_dimension_chart_path, String cc_template, String cc_xls_path, String invoiceNumber,
			String etd, String eta) {
		this.error = "";
		this.pi_path = pi_path;
		this.oceanBillOfLading_path = oceanBillOfLading_path;
		this.product_dimension_chart_path = product_dimension_chart_path;
		this.cc_template = cc_template;
		this.cc_xls_path = cc_xls_path;
		this.invoiceNumber = invoiceNumber;
		this.eta = eta;
		this.etd = etd;

	}
	
	
	public String run() throws IOException {
		
		this.workbook =  WorkbookFactory.create(new File(cc_template));
		
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
		int numOfContainers = pi.getNumberOfContainer();
		if (poNumber != null) {
			cell = worksheet.getRow(PO_NUMBER_ROW).getCell(PO_NUMBER_COL);
			cell.setCellValue(poNumber);
			
			
			

		if (cc_xls_path.trim().isEmpty()) {
			this.cc_xls_path = poNumber + " CI & PL";
        } else {
        	this.cc_xls_path += "/"+poNumber + " CI & PL";
        }
			
			
			// Rename sheets
			this.workbook.setSheetName(0, poNumber + " MASTER CI");
			this.workbook.setSheetName(1, poNumber + "-1 CI");
			this.workbook.setSheetName(2, poNumber + "-1 PL");
			if (numOfContainers > 0) {
				for (int i = 2; i <= numOfContainers; i++) {
					this.workbook.cloneSheet(1);
					this.workbook.setSheetName(i*2-1, poNumber + "-" + i + " CI");
					this.workbook.cloneSheet(2);
					this.workbook.setSheetName(i*2, poNumber + "-" + i + " PL");
				}
			} else {
				error += "ERROR: Container No. does not contain a * sign.\n";
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
		

		
		// Master CI 
		List<Item> items = pi.getItems();

		if (items != null && !items.isEmpty()) {
			for (Item item : items) {
				Util.copyRow(workbook, worksheet, this.item_row, this.item_row + 1);
				worksheet.getRow(this.item_row + 1 ).setHeight(worksheet.getRow(this.item_row).getHeight());
				// Style # (Part No.)
				cell = worksheet.getRow(this.item_row).getCell(STYLE_NUMBER_COL);
				cell.setCellValue(item.getPartNum());
				
				// Vendor Style # (Item #)
				cell = worksheet.getRow(this.item_row).getCell(VENDOR_STYLE_NUMBER_COL);
				cell.setCellValue(item.getItemNum());
				
				// Description
				cell = worksheet.getRow(this.item_row).getCell(DESCRIPTION_COL);
				cell.setCellValue(item.getDescription());
				
				// Material
				cell = worksheet.getRow(this.item_row).getCell(MATERIAL_COL);
				cell.setCellValue("100% Polyester");			
				
				// HTS Code
				cell = worksheet.getRow(this.item_row).getCell(HTS_CODE_COL);
				cell.setCellValue("9401.61.9000");
				
				// QTY 
				cell = worksheet.getRow(this.item_row).getCell(QUANTITY_COL);
				cell.setCellValue(item.getQuantity());
				
				// Carton = QTY
				cell = worksheet.getRow(this.item_row).getCell(CARTON_COL);
				cell.setCellValue(item.getQuantity());
				
				// Package Price 
				cell = worksheet.getRow(this.item_row).getCell(PACKAGE_PRICE_COL);
				cell.setCellValue(0.00);
				
				// Unit Price
				cell = worksheet.getRow(this.item_row).getCell(UNIT_PRICE_COL);
				cell.setCellValue(item.getUnitPrice());
				
				// Total Amount
				cell = worksheet.getRow(this.item_row).getCell(TOTAL_AMOUNT_COL);
				cell.setCellValue(item.getTotalAmount());
				
				this.item_row++;
	
			}
			
			// Removed default 2 empty rows
			worksheet.removeRow(worksheet.getRow(this.item_row));
			worksheet.removeRow(worksheet.getRow(this.item_row+1));
		} else {
			error += "ERROR: Items not found in given PI.\n" + 
					"错误： 找不到任何产品在PI里.\n";
		}
		
		
		
		// CI and PI for each container
		ProductDimensionChart pdc = new ProductDimensionChart(product_dimension_chart_path);
		
		for (int i = 1; i <= numOfContainers; i++) {
			List<Item> containerItems = pdc.getContainerItems(i);
			if (containerItems != null) {
				
				
//--------------PL-------------------------------------------------------------------------------------------------------
				this.item_row = 20;
				worksheet = workbook.getSheet(poNumber + "-" + i + " PL");
				
				// PL - Container Qty
				containerQty = pi.getContainerSize();
				if (containerQty != null) {
					containerQty = "1*" + containerQty;
					cell = worksheet.getRow(CONTAINER_QTY_ROW).getCell(CONTAINER_QTY_COL);
					cell.setCellValue(containerQty);
				}
				
				int totalCarton = 0;
				double totalGrossWeight = 0.0;
				double totalCbm = 0.0;
				
				for (Item containerItem : containerItems) {
					
					Item item = Util.findItem(items, containerItem.getStyleNum(), containerItem.getVendorStyleNum());
					
					Util.copyRow(workbook, worksheet, this.item_row, this.item_row + 1);
					worksheet.getRow(this.item_row + 1 ).setHeight(worksheet.getRow(this.item_row).getHeight());
					// Style # (Part No.)
					cell = worksheet.getRow(this.item_row).getCell(STYLE_NUMBER_COL);
					cell.setCellValue(containerItem.getStyleNum());
					
					// Description
					cell = worksheet.getRow(this.item_row).getCell(DESCRIPTION_COL);
					cell.setCellValue(containerItem.getDescription());
					
					// Vendor Style # (Item #)
					cell = worksheet.getRow(this.item_row).getCell(VENDOR_STYLE_NUMBER_COL);
					cell.setCellValue(containerItem.getVendorStyleNum());
					
					// QTY 
					double quantity = containerItem.getQuantity();
					cell = worksheet.getRow(this.item_row).getCell(QUANTITY_COL);
					cell.setCellValue(quantity);
					
					// Carton = QTY
					totalCarton += quantity;
					cell = worksheet.getRow(this.item_row).getCell(CARTON_COL);
					cell.setCellValue(quantity);
					
					// Material
					cell = worksheet.getRow(this.item_row).getCell(MATERIAL_COL);
					cell.setCellValue("100% Polyester");			
					
					// HTS Code
					cell = worksheet.getRow(this.item_row).getCell(HTS_CODE_COL);
					cell.setCellValue("9401.61.9000");
					
					// Net Weight
					cell = worksheet.getRow(this.item_row).getCell(NET_WEIGHT_COL);
					cell.setCellValue(containerItem.getNetWeight());
						
					// Gross Weight
					double grossWeight = containerItem.getGrossWeight();
					totalGrossWeight += grossWeight;
					cell = worksheet.getRow(this.item_row).getCell(GROSS_WEIGHT_COL);
					cell.setCellValue(grossWeight);
					
					// CMB
					double cbm = containerItem.getCbm();
					totalCbm += cbm;
					cell = worksheet.getRow(this.item_row).getCell(CMB_COL);
					cell.setCellValue(cbm);
					
					this.item_row++;

				}
				// Removed default 2 empty rows
				worksheet.removeRow(worksheet.getRow(this.item_row));
				worksheet.removeRow(worksheet.getRow(this.item_row+1));
				
				// PL - Container #
				String containerNum = oblm.getContainerNumber(totalCarton, Math.round(totalGrossWeight * 100.0) / 100.0 , Math.round(totalCbm * 100.0) / 100.0);
				if (containerNum != null) {
					cell = worksheet.getRow(CONTAINER_NUMBERS_ROW).getCell(CONTAINER_NUMBERS_COL);
					cell.setCellValue(containerNum);
				}
				
				
//--------------CI-------------------------------------------------------------------------------------------------------
				this.item_row = 20;
				worksheet = workbook.getSheet(poNumber + "-" + i + " CI");
				
				// CI - Container Qty
				if (containerQty != null) {
					cell = worksheet.getRow(CONTAINER_QTY_ROW).getCell(CONTAINER_QTY_COL);
					cell.setCellValue(containerQty);
				}
				
				// CI - Container #
				if (containerNum != null) {
					cell = worksheet.getRow(CONTAINER_NUMBERS_ROW).getCell(CONTAINER_NUMBERS_COL);
					cell.setCellValue(containerNum);
				}
				
				for (Item containerItem : containerItems) {
					
					Item item = Util.findItem(items, containerItem.getStyleNum(), containerItem.getVendorStyleNum());
					
					Util.copyRow(workbook, worksheet, this.item_row, this.item_row + 1);
					worksheet.getRow(this.item_row + 1 ).setHeight(worksheet.getRow(this.item_row).getHeight());
					// Style # (Part No.)
					cell = worksheet.getRow(this.item_row).getCell(STYLE_NUMBER_COL);
					cell.setCellValue(containerItem.getStyleNum());
					
					// Description
					cell = worksheet.getRow(this.item_row).getCell(DESCRIPTION_COL);
					cell.setCellValue(containerItem.getDescription());
					
					// Vendor Style # (Item #)
					cell = worksheet.getRow(this.item_row).getCell(VENDOR_STYLE_NUMBER_COL);
					cell.setCellValue(containerItem.getVendorStyleNum());
					
					// QTY 
					double quantity = containerItem.getQuantity();
					cell = worksheet.getRow(this.item_row).getCell(QUANTITY_COL);
					cell.setCellValue(quantity);
					
					// Carton = QTY
					cell = worksheet.getRow(this.item_row).getCell(CARTON_COL);
					cell.setCellValue(quantity);
					
					// Package Price 
					cell = worksheet.getRow(this.item_row).getCell(PACKAGE_PRICE_COL);
					cell.setCellValue(0.00);
					
					// Material
					cell = worksheet.getRow(this.item_row).getCell(MATERIAL_COL);
					cell.setCellValue("100% Polyester");			
					
					// HTS Code
					cell = worksheet.getRow(this.item_row).getCell(HTS_CODE_COL);
					cell.setCellValue("9401.61.9000");

					if (item != null) {
						// Unit Price
						cell = worksheet.getRow(this.item_row).getCell(UNIT_PRICE_COL);
						cell.setCellValue(item.getUnitPrice());
						
						// Total Amount
						cell = worksheet.getRow(this.item_row).getCell(TOTAL_AMOUNT_COL);
						cell.setCellValue(quantity * item.getUnitPrice());
					} else {
						error += "ERROR: Please investigate Style #: "+containerItem.getStyleNum()+", Vendor Style #: "+containerItem.getVendorStyleNum()+".\n" + 
								"错误：请检查Style #: "+containerItem.getStyleNum()+", Vendor Style #: "+containerItem.getVendorStyleNum()+".\n";
					}
					
					this.item_row++;

				}
				// Removed default 2 empty rows
				worksheet.removeRow(worksheet.getRow(this.item_row));
				worksheet.removeRow(worksheet.getRow(this.item_row+1));
				
			} else {
				error += "ERROR: Container measurements not found for container " + i + ".\n" + 
						"错误： 找不到柜"+i+"的分货净毛体信息.\n";
			}
		}
		

		
		
		if (error.isEmpty()) {
			error = "Generated without error.";
		}
		// refreshes all formulas existed in the spreadsheet
//			XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
		HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

		cc_xls_path = Util.correctFileFormat(".xls", cc_xls_path);

		// Open FileOutputStream to write updates
		FileOutputStream output_file;

		try {
			output_file = new FileOutputStream(new File(cc_xls_path));
			// write changes
			workbook.write(output_file);

			// close the stream
			output_file.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		workbook.close();
		
		return error;
	}
}



















