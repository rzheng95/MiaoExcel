package com.rzheng.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ProformaInvoice {
	public static void main(String[] args) {
		ProformaInvoice pi = new ProformaInvoice("052059 PI.pdf");
		List<Item> items = pi.getItems();
		List<Object> stats = pi.getStats(items);
		for (Object s : stats)
		{
			System.out.println(s);
		}
	}

	private String pi_pdf_path;

	public ProformaInvoice(String pi_pdf_path) {
		this.pi_pdf_path = pi_pdf_path;
	}
	
	public String[] readLines() {
		String[] lines = Util.read(this.pi_pdf_path).split("\\r?\\n");
		
		// remove extra spaces
		for (int i = 0; i < lines.length; i++) {
			lines[i] = lines[i].trim().replaceAll(" +", " ");
		}
		
		return lines;
	}
	
	public List<Item> getItems() {
		
		String[] lines = this.readLines();
		List<Item> items = new ArrayList<>();
		int i = 0;
		while (i < lines.length) {

			// U3446-20- Silver Sofa KD / KD sofa argent 9401.61 67.80 24 $314.00 $7,536.00
			if (Util.countString(lines[i], "$", 2) && Util.countString(lines[i], "-", 2)) {
				
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
					for (int j = 2; j < cols.size() - 5; j++) {
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
	
	
	public List<Object> getStats(List<Item> items) {
		
		List<Object> list = new ArrayList<>();
		int totalQuantity = 0;
		double totalNetWeight = 0;
		double totalGrossWeight = 0;
		double totalVolume = 0;
		String errorCode = "";
		
		for (Item item : items) {
			String[] models = null;
			try {
				models = Util.fetchModel(item.getItemNumber());
			} catch (InvalidFormatException | IOException e) {
				e.printStackTrace();
			}
	
			if (models != null) {
	
				double[] stats = null;
				try {
					stats = Util.fetchStats(models[0].substring(0, 6), models[1], Integer.parseInt(item.getQuantity()));
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
							"\n错误： 找不到艺贝型号" + models[0].substring(0, 6) + "\n";
				}
			} else {
				// models is null = itemNum not found
				errorCode = "ERROR: Cannot find item number " + item.getItemNumber() +
						"\n错误： 找不到客户型号" + item.getItemNumber() + "\n";
			}
		}
		list.add(totalQuantity);
		list.add(totalNetWeight);
		list.add(totalGrossWeight);
		list.add(totalVolume);
		list.add(errorCode);
		return list;
	}
	
	
}