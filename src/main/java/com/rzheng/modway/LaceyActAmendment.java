package com.rzheng.modway;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.rzheng.util.Util;

public class LaceyActAmendment {
	public static void main(String[] args) {
		
	}
	
	private String error;
	private String pi_path;
	private String oceanBillOfLading_path;
	private String lacey_act_template;
	private String output_address;
	private String eta;
	
	public LaceyActAmendment(String pi_path, String oceanBillOfLading_path, String lacey_act_template, String output_address, String eta) {
		this.error = "";
		this.pi_path = pi_path;
		this.oceanBillOfLading_path = oceanBillOfLading_path;
		this.lacey_act_template = lacey_act_template;
		this.output_address = output_address;
		this.eta = eta;
	}

	public String run() throws IOException, InvalidFormatException {
		
		XWPFDocument doc = new XWPFDocument(OPCPackage.open(lacey_act_template));
		Util.replaceDocxText(doc, "0008864", "PLEASE WORK!!!");
		doc.write(new FileOutputStream(output_address));
		
		return this.error;
	}
}
