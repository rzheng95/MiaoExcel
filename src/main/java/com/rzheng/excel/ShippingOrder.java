package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import com.rzheng.magnussen.ProformaInvoice;
import com.rzheng.magnussen.ShippingInstructions;
import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class ShippingOrder {
	private final int CONSIGNEE_ROW = 8, CONSIGNEE_COL = 0;

	private final int NOTIFY_ROW = 15, NOTIFY_COL = 0;

	private final int PORT_OF_DISCHARGE_ROW = 23, PORT_OF_DISCHARGE_COL = 0;

	private final int SEA_AIR_ROW = 11, SEA_AIR_COL = 4;

	private final int PORT_OF_LOADING_ROW = 21, PORT_OF_LOADING_COL = 2;

	private final int DESTINATION_ROW = 23, DESTINATION_COL = 2;

	private final int SHIP_TO_ADDRESS_ROW = 35, SHIP_TO_ADDRESS_COL = 1;

	private final int BILL_OF_LADING_REQUIREMENT_ROW = 14, BILL_OF_LADING_REQUIREMENT_COL = 5;

	private final int PO_ROW = 28, PO_COL = 2;

	private final int CPO_ROW = 29, CPO_COL = 2;

	private final int FORWARDER_ROW = 42, FORWARDER_COL = 1;

	private final int x20_ROW = 16, x20_COL = 3;

	private final int x40_ROW = 17, x40_COL = 3;

	private final int x40HC_ROW = 18, x40HC_COL = 3;

	private final int QUANTITY_ROW = 25, QUANTITY_COL = 2;

	private final int GROSS_WEIGHT_ROW = 25, GROSS_WEIGHT_COL = 4;

	private final int MEASUREMENT_ROW = 25, MEASUREMENT_COL = 6;

	// Read the spreadsheet that needs to be updated
	private FileInputStream fsIP;
	// Access the workbook
	private HSSFWorkbook wb;
	// Access the worksheet, so that we can update / modify it.
	private HSSFSheet worksheet;
	// declare a Cell object
	private Cell cell;
	
	private String error;
	private String product_file_path;
	private String dimension_file_path;
	private String si_pdf_path;
	private String pi_pdf_path;
	private String so_xls_path;
	private String si_template;

	public ShippingOrder(String product_file_path, String dimension_file_path, String si_pdf_path, String pi_pdf_path, String so_xls_path, String si_template) {
		this.error = "";
		this.product_file_path = product_file_path;
		this.dimension_file_path = dimension_file_path;
		this.si_pdf_path = si_pdf_path;
		this.pi_pdf_path = pi_pdf_path;
		this.so_xls_path = so_xls_path;
		this.si_template = si_template;
		
		this.cell = null;
		
	}

	public String run() throws IOException  {
		
		try {
			this.fsIP = new FileInputStream(new File(this.si_template));
		} catch (FileNotFoundException e) {
			error = "ERROR: Shipping Order Template Not Found!\n" + this.si_template; 
			e.printStackTrace();
			return error;
		}
		
		try {
			this.wb = new HSSFWorkbook(fsIP);
		} catch (IOException e) {
			error = "ERROR: FileInputStream Exception";
			e.printStackTrace();
			return error;
		}
		
		this.worksheet = wb.getSheetAt(0);
		if (!error.isEmpty()) {
			return error;
		}
		
		ShippingInstructions si = new ShippingInstructions(si_pdf_path);
		
		// Consignee
		String consignee = si.getConsignee();
		if(consignee != null) {
			cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
			cell.setCellValue(consignee);
		} else {
			error += "ERROR: Consignee not found.\n" +
					"错误： 找不到Consignee.\n";
		}
		
		// Notify Party	
		String notifyParty = si.getNotifyParty();
		if(notifyParty != null) {
			cell = worksheet.getRow(NOTIFY_ROW).getCell(NOTIFY_COL);
			cell.setCellValue(notifyParty);
		} else {
			error += "ERROR: Notify Party not found.\n" +
					"错误： 找不到Notify Party.\n";
		}
		
		// Ship-to Address	
		String shipToAddress = si.getShipToAddress();
		if(shipToAddress != null) {
			cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
			cell.setCellValue(shipToAddress);
		} else {
			error += "ERROR: Ship-to Address not found.\n" +
					"错误： 找不到Ship-to Address.\n";
		}
		
		// Forwarder	
		String forwarder = si.getForwarder();
		if(forwarder != null) {
			cell = worksheet.getRow(FORWARDER_ROW).getCell(FORWARDER_COL);
			cell.setCellValue(forwarder);
		} else {
			error += "ERROR: Forwarder not found.\n" +
					"错误： 找不到Forwarder.\n";
		}
		
		// Port of Discharge
		String portOfDischarge = si.getPortOfDischarge();
		if(portOfDischarge != null) {
			cell = worksheet.getRow(PORT_OF_DISCHARGE_ROW).getCell(PORT_OF_DISCHARGE_COL);
			cell.setCellValue(portOfDischarge);
			
			cell = worksheet.getRow(SEA_AIR_ROW).getCell(SEA_AIR_COL);
			if (portOfDischarge.toUpperCase().contains(Constants.SEA))
				cell.setCellValue(Constants.SEA);
			else
				cell.setCellValue(Constants.AIR);
		} else {
			error += "ERROR: Port of Discharge not found.\n" +
					"错误： 找不到Port of Discharge.\n";
		}
		
		// Port of Loading	
		String portOfLoading = si.getPortOfLoading();
		if(portOfLoading != null) {
			cell = worksheet.getRow(PORT_OF_LOADING_ROW).getCell(PORT_OF_LOADING_COL);
			cell.setCellValue(portOfLoading);
		} else {
			error += "ERROR: Port of Loading not found.\n" +
					"错误： 找不到Port of Loading.\n";
		}
		
		// Destination (Place of Delivery)
		String destination = si.getDestination();
		if(destination != null) {
			cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
			cell.setCellValue(destination);
		} else {
			error += "ERROR: Destination not found.\n" +
					"错误： 找不到Destination.\n";
		}
		
		// PO #	
		String poNumber = si.getPoNumber();
		if(poNumber != null) {
			cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
			cell.setCellValue(poNumber);
			String[] arr = poNumber.split(" ");
			if (so_xls_path.trim().isEmpty()) {
				
				if (arr != null && arr.length == 3)
					so_xls_path = "Shipping Order " + arr[2]; // + PO
				else {
					SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");  
				    Date date = new Date();  
					so_xls_path = "Shipping Order " + formatter.format(date);
				}
	        } else {
	        	if (arr != null && arr.length == 3)
	        		so_xls_path += "/Shipping Order " + arr[2];
	        }
		} else {
			error += "ERROR: PO # not found.\n" +
					"错误： 找不到PO #.\n";
		}
		
		// CPO #	
		String cpoNumber = si.getCpoNumber();
		if(cpoNumber != null) {
			cell = worksheet.getRow(CPO_ROW).getCell(CPO_COL);
			cell.setCellValue(cpoNumber);
		} else {
			error += "ERROR: CPO # not found.\n" +
					"错误： 找不到CPO #.\n";
		}
		
		// Container Size
		String containerSize = si.getContainerSize();
		if(containerSize != null) {
			if (containerSize.equalsIgnoreCase(Constants._20)) {
				cell = worksheet.getRow(x20_ROW).getCell(x20_COL);
				cell.setCellValue(1);
			} else if (containerSize.equalsIgnoreCase(Constants._40)) {
				cell = worksheet.getRow(x40_ROW).getCell(x40_COL);
				cell.setCellValue(1);
			} else if (containerSize.equalsIgnoreCase(Constants._40HC)) {
				cell = worksheet.getRow(x40HC_ROW).getCell(x40HC_COL);
				cell.setCellValue(1);
			}
		} else {
			error += "ERROR: Container size not found.\n" +
					"错误： 找不到Container size.\n";
		}
		
		// Bill of Lading Requirement
		String billOfLading = si.getBillOfLadingRequirement();
		if(billOfLading != null) {
			cell = worksheet.getRow(BILL_OF_LADING_REQUIREMENT_ROW).getCell(BILL_OF_LADING_REQUIREMENT_COL);
			cell.setCellValue(billOfLading);
		} else {
			error += "ERROR: Bill of Lading Requirement not found.\n" +
					"错误： 找不到Bill of Lading Requirement.\n";
		}
		
		
		
		/*
		 * PI
		 * 
		*/
		ProformaInvoice pi = new ProformaInvoice(product_file_path, dimension_file_path, pi_pdf_path);
		
		List<Object> stats = pi.getTotalStats(pi.getItems());
		
		if (stats != null) {
			
			if (stats.get(Constants.ERROR_CODE_INDEX).toString().isEmpty()) {
				cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
				cell.setCellValue(Integer.parseInt(stats.get(Constants.TOTAL_QUANTITY_INDEX).toString()));
				
				cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
				cell.setCellValue(Double.parseDouble(stats.get(Constants.TOTAL_GROSS_WEIGHT_INDEX).toString()));
				
				cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
				cell.setCellValue(stats.get(Constants.TOTAL_VOLUME_INDEX).toString() + " CBM");
			} else {
				error = stats.get(Constants.ERROR_CODE_INDEX).toString();
				
				CellStyle backgroundStyle = wb.createCellStyle();
				backgroundStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
				backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
				cell.setCellStyle(backgroundStyle);
				
				cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
				cell.setCellStyle(backgroundStyle);
				
				cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
				cell.setCellStyle(backgroundStyle);
			}
		} else {
			error += "ERROR: No Item Found in the Provided Pro Forma Invoice File,\n"
					+ "Or Incorrect Pro Forma Invoice File Path." 
					+ this.pi_pdf_path + "\n"
					+ "Perhaps You Entered the Incorrect File Path? Zhuzhu!";
		}
		
	
		
		// refreshes all formulas existed in the spreadsheet
		HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
		// Close the InputStream
		if(fsIP != null)
			fsIP.close();

		if (error.isEmpty()) {
			error = "Success!";
			
			so_xls_path = Util.correctXlsFilename(so_xls_path);
			// Open FileOutputStream to write updates
			FileOutputStream output_file = new FileOutputStream(new File(so_xls_path));
			// write changes
			wb.write(output_file);
			// close the stream
			output_file.close();
		} 
		
		return error;
	}

}






