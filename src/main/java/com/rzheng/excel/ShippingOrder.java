package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

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

	private String error = "";
	
	private boolean model_not_found_error = false;

	// Read the spreadsheet that needs to be updated
	FileInputStream fsIP = new FileInputStream(new File("Shipping  Order Template.xls"));
	// Access the workbook
	HSSFWorkbook wb = new HSSFWorkbook(fsIP);
	// Access the worksheet, so that we can update / modify it.
	HSSFSheet worksheet = wb.getSheetAt(0);

	// declare a Cell object
	Cell cell = null;

	public ShippingOrder() throws IOException, InvalidFormatException {

	}

	public String run(String si_pdf_path, String pi_pdf_path, String so_xls_path)
			throws IOException, InvalidFormatException {
		
		ShipmentInformation si = new ShipmentInformation(si_pdf_path);
		
		// Consignee
		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
		String consignee = si.getConsignee();
		if(consignee != null) {
			cell.setCellValue(consignee);
		} else {
			error = "ERROR: Consignee not found.\n" +
					"错误： 找不到Consignee.\n";
		}
		
		// Notify Party
		cell = worksheet.getRow(NOTIFY_ROW).getCell(NOTIFY_COL);
		String notifyParty = si.getNotifyParty();
		if(notifyParty != null) {
			cell.setCellValue(notifyParty);
		} else {
			error = "ERROR: Notify Party not found.\n" +
					"错误： 找不到Notify Party.\n";
		}
		
		// Ship-to Address
		cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
		String shipToAddress = si.getShipToAddress();
		if(shipToAddress != null) {
			cell.setCellValue(shipToAddress);
		} else {
			error = "ERROR: Ship-to Address not found.\n" +
					"错误： 找不到Ship-to Address.\n";
		}
		
		// Forwarder
		cell = worksheet.getRow(FORWARDER_ROW).getCell(FORWARDER_COL);
		String forwarder = si.getForwarder();
		if(forwarder != null) {
			cell.setCellValue(forwarder);
		} else {
			error = "ERROR: Forwarder not found.\n" +
					"错误： 找不到Forwarder.\n";
		}
		
		// Port of Discharge
		cell = worksheet.getRow(PORT_OF_DISCHARGE_ROW).getCell(PORT_OF_DISCHARGE_COL);
		String portOfDischarge = si.getPortOfDischarge();
		if(portOfDischarge != null) {
			cell.setCellValue(portOfDischarge);
			
			cell = worksheet.getRow(SEA_AIR_ROW).getCell(SEA_AIR_COL);
			if (portOfDischarge.toUpperCase().contains(Constants.SEA))
				cell.setCellValue(Constants.SEA);
			else
				cell.setCellValue(Constants.AIR);
		} else {
			error = "ERROR: Port of Discharge not found.\n" +
					"错误： 找不到Port of Discharge.\n";
		}
		
		// Port of Loading
		cell = worksheet.getRow(PORT_OF_LOADING_ROW).getCell(PORT_OF_LOADING_COL);
		String portOfLoading = si.getPortOfLoading();
		if(portOfLoading != null) {
			cell.setCellValue(portOfLoading);
		} else {
			error = "ERROR: Port of Loading not found.\n" +
					"错误： 找不到Port of Loading.\n";
		}
		
		// Destination (Place of Delivery)
		cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
		String destination = si.getDestination();
		if(destination != null) {
			cell.setCellValue(destination);
		} else {
			error = "ERROR: Destination not found.\n" +
					"错误： 找不到Destination.\n";
		}
		
		// PO #
		cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
		String poNumber = si.getPoNumber();
		if(poNumber != null) {
			cell.setCellValue(poNumber);
		} else {
			error = "ERROR: PO # not found.\n" +
					"错误： 找不到PO #.\n";
		}
		
		// CPO #
		cell = worksheet.getRow(CPO_ROW).getCell(CPO_COL);
		String cpoNumber = si.getCpoNumber();
		if(cpoNumber != null) {
			cell.setCellValue(cpoNumber);
		} else {
			error = "ERROR: CPO # not found.\n" +
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
			error = "ERROR: Container size not found.\n" +
					"错误： 找不到Container size.\n";
		}
		
		
		

		// PI
		String pi = Util.read(pi_pdf_path);

		String[] lines = pi.split("\\r?\\n");
		int i = 0;

		// remove extra spaces
		while (i < lines.length) {
			lines[i] = lines[i].trim().replaceAll(" +", " ");
			i++;
		}

		i = 0;

//		double totalNetWeight = 0;
		double totalGrossWeight = 0;
		double totalVolume = 0;
		int totalQuantity = 0;

		while (i < lines.length) {

			// U3446-20- Silver Sofa KD / KD sofa argent 9401.61 67.80 24 $314.00 $7,536.00
			if (Util.countString(lines[i], "$", 2) && Util.countString(lines[i], "-", 2)) {

				String item = lines[i];
				String[] arr = item.split(" ");

				String itemNum = arr[0];
				int quantity = Integer.parseInt(arr[arr.length - 3]);
				totalQuantity += quantity;

				i++;
				arr = lines[i].split(" ");
				itemNum += arr[0];

				String[] models = Util.fetchModel(itemNum);

				if (models != null) {

					double[] stats = Util.fetchStats(models[0].substring(0, 6), models[1], quantity);
					
					if (stats != null && stats.length == 3) {
//						totalNetWeight += stats[0];
						totalGrossWeight += stats[1];
						totalVolume += stats[2];
					} else {
						this.model_not_found_error = true;
						this.error += "ERROR: model number " + models[0].substring(0, 6) + " cannot be found\n"
								+ "错误： 找不到艺贝型号" + models[0].substring(0, 6) + "\n";
					}

				} else {
					
					this.error += "ERROR: item number " + itemNum + " cannot be found\n"
							+ "错误： 找不到客户型号" + itemNum + "\n";
				}

			}


			i++;
		}
		
		if(!this.model_not_found_error) {
			cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
			cell.setCellValue(totalGrossWeight);

			cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
			cell.setCellValue(totalVolume + " CBM");	
		} else {
			CellStyle backgroundStyle = wb.createCellStyle();
			backgroundStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); 
			cell = worksheet.getRow(GROSS_WEIGHT_ROW).getCell(GROSS_WEIGHT_COL);
			cell.setCellStyle(backgroundStyle);
			
			cell = worksheet.getRow(MEASUREMENT_ROW).getCell(MEASUREMENT_COL);
			cell.setCellStyle(backgroundStyle);
		}
		
		cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
		cell.setCellValue(totalQuantity);
		

		// Close the InputStream
		fsIP.close();

		if (!so_xls_path.contains(".xls") && !so_xls_path.isEmpty())
			so_xls_path = so_xls_path + ".xls";
		if (so_xls_path.contains(".xlsx"))
			so_xls_path = so_xls_path.substring(0, so_xls_path.length() - 1);

		// Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File(so_xls_path));
		// write changes
		wb.write(output_file);
		// close the stream
		output_file.close();

		return error;
	}

}
