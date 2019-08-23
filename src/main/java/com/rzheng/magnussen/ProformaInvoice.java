package com.rzheng.magnussen;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class ProformaInvoice {
	public static void main(String[] args) {
		ProformaInvoice pi = new ProformaInvoice("magnussen/magnussen 产品对照表 201905025.xlsx", "magnussen/净毛体统计2016.09.07.xls", "magnussen/49325/049325 PI.pdf");
		for (Item i : pi.getItems()) {
			//System.out.println(i.getDescription());
		}
	}

	private String pi_pdf_path;
	private String product_file_path;
	private String dimension_file_path;

	public ProformaInvoice(String product_file_path, String dimension_file_path, String pi_pdf_path) {
		this.pi_pdf_path = pi_pdf_path;
		this.product_file_path = product_file_path;
		this.dimension_file_path = dimension_file_path;
	}
	
	public String[] readLines() {
		String text = Util.readPDF(this.pi_pdf_path);
		if (text == null)
			return null;
		
		String[] lines = text.split("\\r?\\n");
		
		// remove extra spaces
		for (int i = 0; i < lines.length; i++) {
			lines[i] = lines[i].trim().replaceAll(" +", " ");
		}
		
		return lines;
	}
	
	public List<Item> getItems() {
		
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		List<Item> items = new ArrayList<>();
		int i = 0;
		while (i < lines.length) {

			// U3446-20- Silver Sofa KD / KD sofa argent 9401.61 67.80 24 $314.00 $7,536.00
			if (Util.countString(lines[i], "$", 2) && Util.countString(lines[i], "-", 2)) {
				System.out.println(lines[i]);
				
				List<String> cols = Arrays.asList(lines[i].split(" "));
				cols = new ArrayList<>(cols);


				i++;
				String [] arr = lines[i].split(" ");
				if (arr != null) {
					if (arr.length >= 1) {
						cols.set(0, cols.get(0) + arr[0]);
					}
					if (arr.length >= 2) {
						cols.add(cols.size()-5, arr[1]);
					}
				}
				String description = "";
				if (cols.size() >= 7) {
					for (int j = 1; j < cols.size() - 5; j++) {
						description += cols.get(j);
						if (j < cols.size() - 6)
							description += " ";
					}
				}
				
			
				items.add(new Item(cols.get(0), 
						"", 
						description, 
						cols.get(cols.size()-5),
						cols.get(cols.size()-4), 
						cols.get(cols.size()-3), 
						cols.get(cols.size()-2), 
						cols.get(cols.size()-1)
						));
			}

			i++;
		}
		return items;
	}
	
	
	public List<Object> getTotalStats(List<Item> items) {
		
		if (items == null || items.isEmpty()) {
			return null;
		}
		
		List<Object> list = new ArrayList<>();
		int totalQuantity = 0;
		double totalNetWeight = 0;
		double totalGrossWeight = 0;
		double totalVolume = 0;
		String errorCode = "";
		
		for (Item item : items) {
			String[] models = null;
			try {
				models = Util.fetchModel(this.product_file_path, item.getItemNumber());
			} catch (InvalidFormatException | IOException e) {
				e.printStackTrace();
			}
	
			if (models != null) {
	
				double[] stats = null;
				try {
					stats = Util.fetchDimensions(this.dimension_file_path, models[0].substring(0, 6), models[1], Integer.parseInt(item.getQuantity()));
				} catch (IOException e) {
					e.printStackTrace();
				}
	
				if (stats != null && stats.length == 3) {
					totalQuantity += Integer.parseInt(item.getQuantity());
					totalNetWeight += stats[0];
					totalGrossWeight += stats[1];
					totalVolume += stats[2];
				} else {
					// stats is null = model number not found
					errorCode = "ERROR: Cannot find model number " + models[0].substring(0, 6) +
							"\n错误： 找不到艺贝型号" + models[0].substring(0, 6) + "\n"
									+ "Please check your Dimemsion Chart File.\n"
									+ "请检查净毛体统计表.\n";
				}
			} else {
				// models is null = itemNum not found
				errorCode = "ERROR: Cannot find item number " + item.getItemNumber() +
						"\n错误： 找不到客户型号" + item.getItemNumber() + "\n"
								+ "Please check your Product Chart File.\n"
								+ "请检查产品对照表.\n";
			}
		}

		list.add(totalQuantity);
		list.add(totalNetWeight);
		list.add(totalGrossWeight);
		list.add(totalVolume);
		list.add(errorCode);
		return list;
	}
	
	public double[] getStats(Item item) {
		if (item == null)
			return null;
		
		String[] models = null;
		try {
			models = Util.fetchModel(this.product_file_path, item.getItemNumber());
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		}

		if (models != null) {

			double[] stats = null;
			try {
				stats = Util.fetchDimensions(this.dimension_file_path, models[0].substring(0, 6), models[1], Integer.parseInt(item.getQuantity()));
			} catch (IOException e) {
				e.printStackTrace();
			}

			if (stats != null && stats.length == 3) {
				return new double[] { Integer.parseInt(item.getQuantity()), stats[0], stats[1], stats[2], 1 };

			} else {
				// stats is null = model number not found
				return new double[] { -2 };
			}
		} else {
			// models is null = itemNum not found
			return new double[] { -1 };
		}
		
	}
	
	public double getTotalExclTaxAmount() {
		String[] lines = this.readLines();
		if (lines == null)
			return -1;
		
		int i = 0;
		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.TOTAL)) {
				
				if (lines[i].toUpperCase().contains(Constants.TOTAL_EXCL_TAX)) {
					
					String[] arr = lines[i].split(" ");
					String amountStr = arr[arr.length-1];
					if (amountStr != null && !amountStr.isEmpty())
						return Util.extractNumberFromAmount(amountStr);
            		
//            		cell.setCellType(CellType.NUMERIC);
//            		CellStyle cs = wb.createCellStyle();
//            		cs.setDataFormat((short)7);
//            		cell.setCellStyle(cs);
				} 
			}
			i++;
		}
		
		return -1;
	}
	
	public int getQuantity() {
		String[] lines = this.readLines();
		if (lines == null)
			return -1;
		
		int i = 0;
		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.TOTAL)) {
				if (lines[i].toUpperCase().contains(Constants.SUB_TOTAL) ||
						lines[i].toUpperCase().contains(Constants.TOTAL_EXCL_TAX)) {
					i++;
					continue;
				}
				String[] arr = lines[i].split(" ");
				if (arr != null && arr.length >= 3)
					return Integer.parseInt(arr[arr.length-1]);
			}
			i++;
		}
		return -1;
	}
	
}
