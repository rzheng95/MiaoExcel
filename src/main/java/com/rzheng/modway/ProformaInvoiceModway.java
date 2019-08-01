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
		new ProformaInvoiceModway("modway/9395/0009395-PI-MODWAY-041919(1).xls").getContainerQty();
	}
	
	private String pi_path;
	
	// for any Excel version both .xls and .xlsx
	private Workbook wb;
	private Sheet worksheet;
	
	public ProformaInvoiceModway(String pi_path) {		
		this.pi_path = pi_path;
			
	}
	
	public String getContainerQty() throws IOException {

		// for any Excel version both .xls and .xlsx
		Workbook wb = WorkbookFactory.create(new File(pi_path));
		Sheet worksheet = wb.getSheetAt(0);

		if (worksheet != null)
			for (Row row : worksheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getRichStringCellValue().getString().trim().equals(Constants.CONTAINER_NO)) {
							
							wb.close();
							return worksheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex()+1).toString();
						}
					}

				}
			}

		if (wb != null) {
			wb.close();
		}
		return null;
	}
	
}













