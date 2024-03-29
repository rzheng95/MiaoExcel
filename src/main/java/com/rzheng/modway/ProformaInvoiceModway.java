package com.rzheng.modway;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import com.rzheng.util.Constants;

public class ProformaInvoiceModway {
	public static void main(String[] args) throws IOException {

		ProformaInvoiceModway pi = new ProformaInvoiceModway("modway/8864/0008864-PI-MODWAY-121818.xls");

		System.out.println(pi.getPortName());

	}
	
	private String pi_path;
	// for any Excel version both .xls and .xlsx
	private Workbook wb;
	private Sheet worksheet;
	
	public ProformaInvoiceModway(String pi_path) {		
		this.pi_path = pi_path;
		
		try {
			wb = WorkbookFactory.create(new File(pi_path));
			worksheet = wb.getSheetAt(0);
			wb.close();
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}	
		
	}
	
	public String getContainerQty() {

		if (worksheet != null)
			for (Row row : worksheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(Constants.CONTAINER_NO)) {

							return worksheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex() + 1).toString();
						}
					}

				}
			}

		return null;
	}
	
	public int getNumberOfContainer() {
		
		String containerQty = getContainerQty();
		if (containerQty != null) {
			if (containerQty.contains("*")) {
				String[] arr = containerQty.split("\\*");
				if (arr != null && arr.length >= 1) {
					return Integer.parseInt(arr[0].trim());
				}
			}
		}
		
		return -1;
	}
	
	public String getContainerSize() {
		
		String containerQty = getContainerQty();
		if (containerQty != null) {
			if (containerQty.contains("*")) {
				String[] arr = containerQty.split("\\*");
				if (arr != null && arr.length >= 2) {
					return arr[1].trim();
				}
			}
		}
		
		return null;
	}
	
	public String getPoNumber() {

		if (worksheet != null)
			for (Row row : worksheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getRichStringCellValue().getString().trim()
								.equalsIgnoreCase(Constants.PURCHASE_ORDER_NO)) {
							return worksheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex() + 2).toString();
						}
					}

				}
			}

		return null;
	}

	public List<Item> getItems() {

		if (worksheet != null) {
			
			boolean isValid = false;
			List<Item> items = new ArrayList<>();
			
			for (Row row : worksheet) {
				
				if (isValid) {
					if (worksheet.getRow(row.getRowNum()).getCell(2).toString().trim().isEmpty()) {
						return items;
					}
					Item item = new Item(
							worksheet.getRow(row.getRowNum()).getCell(2).toString().trim(), // part No.
							worksheet.getRow(row.getRowNum()).getCell(3).toString().trim(), // Description
							worksheet.getRow(row.getRowNum()).getCell(4).toString().trim(), // Item #
							worksheet.getRow(row.getRowNum()).getCell(5).toString().trim(), // Fabric
							(int) worksheet.getRow(row.getRowNum()).getCell(6).getNumericCellValue(), // Quantity
							worksheet.getRow(row.getRowNum()).getCell(7).getNumericCellValue(), // Unit Price
							worksheet.getRow(row.getRowNum()).getCell(8).getNumericCellValue()); // Total

					items.add(item);

				}
				
				if (!isValid) {
					for (Cell cell : row) {
						if (cell.getCellType() == CellType.STRING) {

							if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(Constants.PART_NO)) {
								isValid = true;
								break;
							}
						}
					}
				}
				
			}
		}
		return null;
	}
	
	public double getTotalAmount() {

		double totalAmount = 0;

		for (Item item : getItems()) {
			totalAmount += item.getTotalAmount();
		}
		
		return totalAmount;
	}
	
	public int getTotalQuantity() {

		int totalQuantity = 0;
		
		for (Item item : getItems()) {
			totalQuantity += item.getQuantity();
		}
	
		return totalQuantity;
	}
	
	public String getShipTo() {
		
		String shipTo = "";
		boolean found = false;
		if (worksheet != null)
			for (Row row : worksheet) {
				
		
				Cell cell = row.getCell(0);
				if (cell != null && cell.getCellType() == CellType.STRING) {
					
					if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(Constants.SHIP_TO)) {
						found = true;
					}
					
					
					
					if (found) {
						if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(Constants.SHIP_VIA)) {
							return shipTo;
						}
					}
				}
				
				if (found) {
					shipTo += row.getCell(1).toString() + "\n";
				}
				
			}

		return null;
	}
	
	
	public String getPortName() {
		String shipTo = getShipTo();
		
		if (shipTo != null) {
			String[] arr = shipTo.split("\\r?\\n");
			
			if (arr != null && arr.length >= 2) {
				String portLine = arr[arr.length-2];
				if (portLine != null && portLine.contains(",")) {
					arr = portLine.split(",");
					if (arr != null && arr.length >= 1)
						return arr[0].trim();
				}
			}
		}
		
		return null;
	}
	
}














