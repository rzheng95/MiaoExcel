package com.rzheng.logistics.companies;

import java.util.Calendar;

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class UnitexLogistics {

	public static void main(String[] args) {
		UnitexLogistics ul = new UnitexLogistics("指代/Unitex Logistics Ltd 051490.docx");

		System.out.println(ul.getMasterBLNumber());
	}
	
	
	private String doc_path;
	
	public UnitexLogistics(String doc_path) {
		this.doc_path = doc_path;
	}
	
	public String getMasterBLNumber() {
		String text = Util.readDocument(this.doc_path);
		if (text != null) {
			String[] lines = text.split("\\r?\\n");
			
			int i = 0;
			while (i < lines.length) {
				if (lines[i].toUpperCase().contains(Constants.MBL)) {
					return lines[i].substring(Constants.MBL.length()).trim();
				}				
				i++;
			}
		}	
		return text;
	}
	
	public String getETD() {
		String text = Util.readDocument(this.doc_path);
		if (text != null) {
			String[] lines = text.split("\\r?\\n");
			
			int i = 0;
			while (i < lines.length) {
				if (lines[i].toUpperCase().contains(Constants.ETD)) {
					String line = lines[i].substring(Constants.ETD.length()).trim();
					if (line.contains(".")) {
						String[] arr = line.split("\\.");				
						if (arr.length == 2) {
							String month = arr[0];
							String day = arr[1];
							Calendar calendar = Calendar.getInstance();
							return month + "/" + day + "/" + calendar.get(Calendar.YEAR);
						}
					}
					return line;
				}				
				i++;
			}
		}	
		return text;
	}

	public String getETA() {
		String text = Util.readDocument(this.doc_path);
		if (text != null) {
			String[] lines = text.split("\\r?\\n");
			
			int i = 0;
			while (i < lines.length) {
				if (lines[i].toUpperCase().contains(Constants.ETA)) {
					String line = lines[i].substring(Constants.ETA.length()).trim();
					if (line.contains(".")) {
						String[] arr = line.split("\\.");				
						if (arr.length == 2) {
							String month = arr[0];
							String day = arr[1];
							Calendar calendar = Calendar.getInstance();
							return month + "/" + day + "/" + calendar.get(Calendar.YEAR);
						}
					}
					return line;
				}				
				i++;
			}
		}	
		return text;
	}

	public String getMotherVessel() {
		String text = Util.readDocument(this.doc_path);
		if (text != null) {
			String[] lines = text.split("\\r?\\n");
			
			int i = 0;
			while (i < lines.length) {
				if (lines[i].toUpperCase().contains(Constants.VV)) {
					return lines[i].substring(Constants.VV.length()).trim();
				}				
				i++;
			}
		}	
		return text;
	}
	
}






















