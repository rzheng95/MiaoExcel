package com.rzheng.magnussen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.rzheng.logistics.companies.UnitexLogistics;
import com.rzheng.util.Constants;
import com.rzheng.util.Item;
import com.rzheng.util.Util;

public class CustomsClearance 
{
	// Factory Invoice - RDC & ADC
	private final int SHIP_TO_NAME_ROW = 15, SHIP_TO_NAME_COL = 1;
	private final int SHIP_TO_ADDRESS_ROW = 16, SHIP_TO_ADDRESS_COL = 1;
	private final int FORWARDER_NAME_ROW = 21, FORWARDER_NAME_COL = 1;
	private final int CARRIER_ROW = 22, CARRIER_COL = 1;
	private final int CONTAINER_NUMBER_ROW = 23, CONTAINER_NUMBER_COL = 1;
	private final int CONTAINER_SIZE_ROW = 24, CONTAINER_SIZE_COL = 1;
	private final int SEAL_NUMBER_ROW = 25, SEAL_NUMBER_COL = 1;
	private final int HOUSE_BL_NUMBER_ROW = 27, HOUSE_BL_NUMBER_COL = 1;
	private final int MASTER_BL_NUMBER_ROW = 28, MASTER_BL_NUMBER_COL = 1;
	
	
	private final int DATE_ROW = 4, DATE_COL = 6;
	private final int INVOICE_NUMBER_ROW = 5, INVOICE_NUMBER_COL = 6;
	private final int DELIVERY_NUMBER_ROW = 6, DELIVERY_NUMBER_COL = 6;
	private final int CPO_ROW = 7, CPO_COL = 6;
	private final int ETD_ROW = 8, ETD_COL = 6;
	private final int CUT_OFF_DATE_ROW = 9, CUT_OFF_DATE_COL = 6;
	private final int ETA_ROW = 10, ETA_COL = 6;
	private final int MOTHER_VESSEL_ROW = 12, MOTHER_VESSEL_COL = 6;
	private final int PORT_OF_ORIGIN_ROW = 13, PORT_OF_ORIGIN_COL = 6;
	private final int PORT_OF_DISCHARGE_ROW = 14, PORT_OF_DISCHARGE_COL = 6;
	private final int DESTINATION_ROW = 15, DESTINATION_COL = 6;
	
	private int item_start_row = 31;
	private final int ITEM_COL = 0;
	private final int DESCRIPTION_COL = 1;
	private final int HTS_CODE_COL = 2;
	private final int COUNTRY_OF_ORIGIN_COL = 4;
	private final int QUANTITY_COL = 5;
	private final int UNIT_PRICE_COL = 6;
	private final int TOTAL_AMOUNT_COL = 7;
	
	// Factory PL revised page (page 3)
	private final int GROSS_WEIGHT_COL = 5;
	private final int VOLUME_COL = 6;
	
	// Read the spreadsheet that needs to be updated
	private FileInputStream fileInput;
	// Access the workbook
	private XSSFWorkbook workbook;
	// Access the worksheet, so that we can update / modify it.
	private XSSFSheet worksheet;
	
	// declare a Cell object
	private Cell cell;
	
	private String error;
	private String product_file_path;
	private String dimension_file_path;
	private String si_pdf_path;
	private String pi_pdf_path;
	private String cc_xlsx_path;
	private String cc_template;
	private String logistics_company;
	private String logistics_confirmation_path;
	private String invoiceNumber;
	private String containerNumber;
	private String sealNumber;

	public CustomsClearance(String product_file_path, String dimension_file_path, String si_pdf_path, String pi_pdf_path, 
			String logistics_company, String logistics_comfirmation_path,
			String cc_xlsx_path, String cc_template, String invoiceNumber, String containerNumber, String sealNumber) throws IOException {
		this.error = "";
		this.product_file_path = product_file_path;
		this.dimension_file_path = dimension_file_path;
		this.si_pdf_path = si_pdf_path;
		this.pi_pdf_path = pi_pdf_path;
		this.cc_xlsx_path = cc_xlsx_path;
		this.cc_template = cc_template;
		this.logistics_company = logistics_company;
		this.logistics_confirmation_path = logistics_comfirmation_path;
		this.invoiceNumber = invoiceNumber;
		this.containerNumber = containerNumber;
		this.sealNumber = sealNumber;
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
			this.workbook = new XSSFWorkbook(fileInput);
		} catch (IOException e) {
			error = "ERROR: FileInputStream Exception";
			e.printStackTrace();
			return error;
		}
		
		
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date current_date = new Date();
        Calendar calendar = Calendar.getInstance();

        String today = dateFormat.format(current_date);

        


		
        
//----- Factory Invoice - RDC & ADC page------------------------------------------------------------------
        worksheet = workbook.getSheetAt(0);
        
        // Date
        cell = worksheet.getRow(DATE_ROW).getCell(DATE_COL);
		cell.setCellValue(today);
		

        if (invoiceNumber.trim().isEmpty()) {
        	invoiceNumber = "INYB" + calendar.get(Calendar.YEAR) + "US" + calendar.get(Calendar.MONTH) + calendar.get(Calendar.DAY_OF_MONTH);
        }

        // Invoice number
        cell = worksheet.getRow(INVOICE_NUMBER_ROW).getCell(INVOICE_NUMBER_COL);
		cell.setCellValue(invoiceNumber);
           
		
		ShippingInstructions si = new ShippingInstructions(si_pdf_path);
		
		
        // Ship-to Name (Consignee)
        String consignee = si.getConsignee();
		if (consignee != null) {
			cell = worksheet.getRow(SHIP_TO_NAME_ROW).getCell(SHIP_TO_NAME_COL);
			cell.setCellValue(consignee);
		} else {
			error += "ERROR: Consignee not found.\n" + 
					"错误： 找不到Consignee.\n";
		}
		
		// Ship-to Address
		String shipToAddress = si.getShipToAddress();
		if (shipToAddress != null) {
			cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
			cell.setCellValue(shipToAddress);
		} else {
			error += "ERROR: Ship-to Address not found.\n" + 
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
			error += "ERROR: Forwarder Name not found.\n" + 
					"错误： 找不到Forwarder Name.\n";
		}
		
		// Carrier Name
		String carrier = si.getCarrier();
		if (carrier != null) {
			cell = worksheet.getRow(CARRIER_ROW).getCell(CARRIER_COL);
			cell.setCellValue(carrier);

		} else {
			error += "ERROR: Carrier not found.\n" + 
					"错误： 找不到Carrier.\n";
		}
		
		// Container Number
		if (this.containerNumber != null && !this.containerNumber.trim().isEmpty()) {
			cell = worksheet.getRow(CONTAINER_NUMBER_ROW).getCell(CONTAINER_NUMBER_COL);
			cell.setCellValue(this.containerNumber);
		}
		
		
		
		// Container Size
		String containerSize = si.getContainerSize();
		if (containerSize != null) {
			cell = worksheet.getRow(CONTAINER_SIZE_ROW).getCell(CONTAINER_SIZE_COL);
			cell.setCellValue(containerSize);
		} else {
			error += "ERROR: Container Size not found.\n" + 
					"错误： 找不到Container Size.\n";
		}
		
		
		// Seal Number
		if (this.sealNumber != null && !this.sealNumber.trim().isEmpty()) {
			cell = worksheet.getRow(SEAL_NUMBER_ROW).getCell(SEAL_NUMBER_COL);
			cell.setCellValue(this.sealNumber);
		}
				
		
		// AMS Bill #?

		
		// House B/L Number
		String houseBLNumber = null;
		if (this.logistics_company.equalsIgnoreCase(Constants.UNITEX_LOGISTICS)) {
			houseBLNumber = new UnitexLogistics(this.logistics_confirmation_path).getMasterBLNumber();
		}

		if (houseBLNumber != null) {
			cell = worksheet.getRow(HOUSE_BL_NUMBER_ROW).getCell(HOUSE_BL_NUMBER_COL);
			cell.setCellValue(houseBLNumber);
			
			cell = worksheet.getRow(MASTER_BL_NUMBER_ROW).getCell(MASTER_BL_NUMBER_COL);
			cell.setCellValue(houseBLNumber);
		} else {
			error += "ERROR: Master B/L Number (MBL) not found.\n" + 
					"错误： 找不到 Master B/L Number (MBL).\n";
		}
				
   
//----- Right Column------------------------------------------------------------------
		
		// Delivery Number (PO #)
		String poNumber = si.getPoNumber();
		if(poNumber != null) {
			cell = worksheet.getRow(DELIVERY_NUMBER_ROW).getCell(DELIVERY_NUMBER_COL);
			cell.setCellValue(poNumber.substring(Constants.PO.length()).trim());
		} else {
			error += "ERROR: Delivery Number (PO #) not found.\n" +
					"错误： 找不到Delivery Number (PO #).\n";
		}
        
		// Customer PO# (CPO #)	
		String cpoNumber = si.getCpoNumber();
		if(cpoNumber != null) {
			cell = worksheet.getRow(CPO_ROW).getCell(CPO_COL);
			cell.setCellValue(cpoNumber.substring(Constants.CPO.length()).replaceAll(":", "").trim());
		} else {
			error += "ERROR: CPO # not found.\n" +
					"错误： 找不到CPO #.\n";
		}
		
		// ETD
		String etd = null;
		if (this.logistics_company.equalsIgnoreCase(Constants.UNITEX_LOGISTICS)) {
			etd = new UnitexLogistics(this.logistics_confirmation_path).getETD();
		}
		
		if (etd != null) {
			cell = worksheet.getRow(ETD_ROW).getCell(ETD_COL);
			try {
				cell.setCellValue(dateFormat.format(dateFormat.parse(etd)));
			} catch (ParseException e1) {
				e1.printStackTrace();
			}
			
			// Cut-off Date (One day before ETD)
			try {
				Date date = dateFormat.parse(etd);
				calendar.setTime(date);
		        calendar.add(Calendar.DAY_OF_MONTH, -1);
		        date = calendar.getTime();
		        cell = worksheet.getRow(CUT_OFF_DATE_ROW).getCell(CUT_OFF_DATE_COL);
				cell.setCellValue(dateFormat.format(date));
			} catch (ParseException e) {
				e.printStackTrace();
			}
		} else {
			error += "ERROR: ETD not found.\n" + 
					"错误： 找不到ETD.\n";
		}
		

		// ETA
		String eta = null;
		if (this.logistics_company.equalsIgnoreCase(Constants.UNITEX_LOGISTICS)) {
			eta = new UnitexLogistics(this.logistics_confirmation_path).getETA();
		}
		
		if (eta != null) {
			cell = worksheet.getRow(ETA_ROW).getCell(ETA_COL);
			try {
				cell.setCellValue(dateFormat.format(dateFormat.parse(eta)));
			} catch (ParseException e) {
				e.printStackTrace();
			}

		} else {
			error+= "ERROR: ETA not found.\n" + 
					"错误： 找不到ETA.\n";
		}
		
		
		// Feeder Vessel?
		
		
		
		
		
		
		
		// Mother Vessel
		String motherVessel = null;
		if (this.logistics_company.equalsIgnoreCase(Constants.UNITEX_LOGISTICS)) {
			motherVessel = new UnitexLogistics(this.logistics_confirmation_path).getMotherVessel();
		}
		
		if (motherVessel != null) {
			cell = worksheet.getRow(MOTHER_VESSEL_ROW).getCell(MOTHER_VESSEL_COL);
			cell.setCellValue(motherVessel);

		} else {
			error += "ERROR: Mother Vessel (V/V) not found.\n" + 
					"错误： 找不到 Mother Vessel (V/V)\n";
		}
		
		
		// Port of Origin
		String portOfOrigin = si.getPortOfLoading();
		if(portOfOrigin != null) {
			cell = worksheet.getRow(PORT_OF_ORIGIN_ROW).getCell(PORT_OF_ORIGIN_COL);
			cell.setCellValue(portOfOrigin);
		} else {
			error += "ERROR: Port of Loading (Origin) not found.\n" +
					"错误： 找不到Port of Loading (Origin).\n";
		}
		
		
		// Port of Discharge
		String portOfDischarge = si.getPortOfDischarge();
		if(portOfDischarge != null) {
			cell = worksheet.getRow(PORT_OF_DISCHARGE_ROW).getCell(PORT_OF_DISCHARGE_COL);
			cell.setCellValue(portOfDischarge);
		} else {
			error += "ERROR: Port of Discharge not found.\n" +
					"错误： 找不到Port of Discharge.\n";
		}
		
		// Final Destination
		String destination = si.getDestination();
		if(destination != null) {
			cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
			cell.setCellValue(destination);
		} else {
			error += "ERROR: Destination not found.\n" +
					"错误： 找不到Destination.\n";
		}
		
		
		// PI
		ProformaInvoice pi = new ProformaInvoice(product_file_path, dimension_file_path, pi_pdf_path);
		
		// Items
		List<Item> items = pi.getItems();
		if (items != null && !items.isEmpty()) {
			
			
			
			for (Item item : items) {
				worksheet = workbook.getSheetAt(0);
				
				cell = worksheet.getRow(item_start_row).getCell(ITEM_COL);
				cell.setCellValue(item.getItemNumber());

				cell = worksheet.getRow(item_start_row).getCell(DESCRIPTION_COL);
				cell.setCellValue(item.getDescription());

				cell = worksheet.getRow(item_start_row).getCell(HTS_CODE_COL);
				cell.setCellValue(Double.parseDouble(item.getHtsCode()));

				String countryOfOrigin = si.getPortOfLoadingCountry();
				if (countryOfOrigin != null) {
					cell = worksheet.getRow(item_start_row).getCell(COUNTRY_OF_ORIGIN_COL);
					cell.setCellValue(countryOfOrigin);
				}

				cell = worksheet.getRow(item_start_row).getCell(QUANTITY_COL);
				cell.setCellValue(Double.parseDouble(item.getQuantity()));

				cell = worksheet.getRow(item_start_row).getCell(UNIT_PRICE_COL);
				cell.setCellValue(Util.extractNumberFromAmount(item.getUnitCost()));

				cell = worksheet.getRow(item_start_row).getCell(TOTAL_AMOUNT_COL);
				cell.setCellValue(Util.extractNumberFromAmount(item.getNetAmount()));
				
				
				// Factory PL revised Page (page 3)
				// G.W. (Kgs) (Gross Weight)
				double[] stats = pi.getStats(item);
				if (stats != null) {
					if (stats[stats.length-1] == 1 && stats.length >= 5) {
						worksheet = workbook.getSheetAt(2);
						// Gross Weight
						cell = worksheet.getRow(item_start_row).getCell(GROSS_WEIGHT_COL);
						cell.setCellValue(stats[2]);

						// Volume
						cell = worksheet.getRow(item_start_row).getCell(VOLUME_COL);
						cell.setCellValue(stats[3]);
					} else {
						if (stats[stats.length-1] == -1)
							error += "ERROR: Item Number not found.\n";
						if (stats[stats.length-1] == -2)
							error += "ERROR: Model Number not found.\n";
					}
				}

				item_start_row++;
			}
				
		} else {
			error += "ERROR: Cannot get stats.\n";
		}
		
    
        
		if (cc_xlsx_path.trim().isEmpty()) {
			String[] poNum = si.getPoNumber().split(" ");
			if (poNum != null && poNum.length >= 3)
				cc_xlsx_path = "Magnussen CI&PL&7 Point " + poNum[2]; // + PO
        }
		
		
		// Close the InputStream
		if(fileInput != null)
			fileInput.close();
		


		
		error = "Success!";

		// refreshes all formulas existed in the spreadsheet
		// HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
		XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

		cc_xlsx_path = Util.correctXlsxFilename(cc_xlsx_path);

		// Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File(cc_xlsx_path));
		// write changes
		workbook.write(output_file);
		// close the stream
		output_file.close();
		

		
		workbook.close();
		
		return this.error;
	}
	
	

}































