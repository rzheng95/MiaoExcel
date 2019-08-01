package com.rzheng.modway;

import com.rzheng.util.Util;

public class OceanBillOfLadingModway {
	public static void main(String[] args) {
		new OceanBillOfLadingModway("modway/9395/0009395 HES19050515-海运提单(代理).PDF").getContainerNumber();
	}
	
	private String oceanBillOfLading_path;
	
	public OceanBillOfLadingModway(String oceanBillOfLading_path) {
		this.oceanBillOfLading_path = oceanBillOfLading_path;
	}
	
	public String getContainerNumber() {
		String text = Util.readPDF(oceanBillOfLading_path);
		
		System.out.println(text);
		return oceanBillOfLading_path;
		
	}
}
