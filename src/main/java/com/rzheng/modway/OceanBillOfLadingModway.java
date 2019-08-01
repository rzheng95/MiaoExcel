package com.rzheng.modway;

import java.util.ArrayList;
import java.util.List;

import com.rzheng.util.Util;

public class OceanBillOfLadingModway {
	public static void main(String[] args) {
		OceanBillOfLadingModway obl = new OceanBillOfLadingModway("modway/9395/0009395 HES19050515-海运提单(代理).PDF");
		System.out.println(obl.getPlaceOfDischarge());
//		System.out.println(Util.readPDF(("modway/9395/0009395 HES19050515-海运提单(代理).PDF")));
	}
	
	private String oceanBillOfLading_path;
	private String[] lines;
	
	public OceanBillOfLadingModway(String oceanBillOfLading_path) {
		this.oceanBillOfLading_path = oceanBillOfLading_path;
		
		String text = Util.readPDF(oceanBillOfLading_path);
		if (text != null) {
			lines = text.split("\\r?\\n");
		}
	}
	
	public List<String> getContainerDescriptions() {
		if (lines != null) {
			int i = 0;
			List<String> list = new ArrayList<>();
			while (i < lines.length) {

				if (Util.countString(lines[i], "/", 5))
					list.add(lines[i].trim());

				i++;
			}
			return list;
		}
		return null;
	}
	
	public String getAllContainerNumbers() {
		List<String> containerDescriptions = this.getContainerDescriptions();
		if (containerDescriptions != null) {
			String containerNumbers = "";
			for (String des : containerDescriptions) {
				String[] arr = des.split("/");
				if (arr != null && arr.length >= 1)
					containerNumbers += arr[0] + "/";
			}
			return containerNumbers.substring(0, containerNumbers.length()-1); // remove the last slash
		}
		
		return null;
	}
	
	
	public String getBillOfLadingNumber() {
		if (lines != null) {

			int i = 0;
			while (i < lines.length) {
				
				if (!lines[i].trim().equals("")) {
					String firstLine = lines[i];
					if (firstLine != null && !firstLine.isEmpty()) {
						String[] arr = firstLine.trim().replaceAll(" +", " ").split(" ");
						if (arr != null && arr[arr.length-1].contains("HES")) {
							return arr[arr.length-1];
						}					
					}
				}
				i++;
			}
		}
		
		return null;
	}
	
	public String getPlaceOfDischarge() {
		
		if (lines != null) {
			int i = 0;
			
			while (i < lines.length) {
				
				if (!lines[i].trim().isEmpty() && Util.countString(lines[i], ",", 2))
				{
					String line = lines[i].trim().replaceAll(" +", " ");
					int index = line.indexOf(",");
//					System.out.println(line);
//					System.out.println(index);
					
					/*
					LONGXUAN VILLAGE,CHONGXIAN TOWN,
					16
					YUHANG,HANGZHOU,CHINA
					6
					138 GEORGES RD DAYTON, NJ 08810 329 WYCKOFF MILLS RD, HIGHTSTOWN, NJ 08520
					21
					NEW YORK,NY NEW YORK,NY JFKDOCS@SHIPALLWAYS.COM
					8
					*/
					
					if (Util.countString(line, line.substring(0, index), 2)) {
						return line.substring(0, index);
					}

				}
				i++;
			}
		}
		
		return null;
	}
	
}




















