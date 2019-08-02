package com.rzheng.modway;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ProductDimensionChart {
	public static void main(String[] args) {
		ProductDimensionChart pdc = new ProductDimensionChart("modway/9395/9395分货-有净毛体(1).xls");
	}
	
	private String product_dimension_chart_path;
	private Workbook wb;
	private Sheet worksheet;
	
	public ProductDimensionChart(String product_dimension_chart_path) {
		this.product_dimension_chart_path = product_dimension_chart_path;
		
		try {
			wb = WorkbookFactory.create(new File(product_dimension_chart_path));
			worksheet = wb.getSheetAt(0);
			wb.close();
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}	
	}

	public int getColumnIndex(String cellValue) {
		if (worksheet != null) {
			for (Row row : worksheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getRichStringCellValue().getString().trim()
								.equalsIgnoreCase(cellValue)) {
							return cell.getColumnIndex();
						}
					}
				}
			}
		}
		return -1;
	}
	
	public List<Item> getContainerItems(int containerNumber) {

		if (worksheet != null) {
			List<Item> items = new ArrayList<>();
			for (Row row : worksheet) {
				
				if (row.getRowNum() > 3) { // item listing starts on the 5th row

					int styleNumCol = 0;
					int descriptionCol = 1;
					int vendorStyleNumCol = 3;
					int netWeightCol = getColumnIndex(("柜" + containerNumber + "净重"));
					int quantityCol = netWeightCol - 1 ;
					int grossWeightCol = netWeightCol + 1 ;
					int volumeCol = grossWeightCol + 1;
					
					// found value
					if (row.getCell(netWeightCol) != null && row.getCell(netWeightCol).getNumericCellValue() != 0 && row.getCell(styleNumCol) != null && row.getCell(vendorStyleNumCol) != null) {
						
						Item item = new Item(row.getCell(styleNumCol).getStringCellValue().trim(), 
								row.getCell(descriptionCol).getStringCellValue().trim(), 
								row.getCell(vendorStyleNumCol).getStringCellValue().trim(), 
								(int) row.getCell(quantityCol).getNumericCellValue(), 
								Math.round(row.getCell(netWeightCol).getNumericCellValue() * 100.0) /100.0, 
								Math.round(row.getCell(grossWeightCol).getNumericCellValue() * 100.0) / 100.0, 
								Math.round(row.getCell(volumeCol).getNumericCellValue() * 100.0) / 100.0);

						items.add(item);
					}
				}
			}
			if (items.isEmpty())
				return null;
			return items;
		}

		return null;
	}
	
}















