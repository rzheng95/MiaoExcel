package com.rzheng.modway;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.rzheng.util.Constants;

public class ProformaInvoiceModway {
	public static void main(String[] args) throws IOException {
		ProformaInvoiceModway pi = new ProformaInvoiceModway("modway/9395/0009395-PI-MODWAY-041919(1).xls");
		System.out.println(pi.getPoNumber());
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
	
	
}













