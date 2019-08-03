package com.rzheng.modway;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import  org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class LaceyActAmendment {
	
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

	public String run() throws IOException, InvalidFormatException, XmlException {
		
		FileInputStream input = new FileInputStream(new File(lacey_act_template));
		XWPFDocument doc = new XWPFDocument(input);
		
		// ETA
		Util.replaceDocxText(doc, Constants.ETA_PLACEHOLDER, this.eta);
	
		OceanBillOfLadingModway obl = new OceanBillOfLadingModway(oceanBillOfLading_path);
		
		// Container Numbers
		String containerNumbers = obl.getAllContainerNumbers();
		if (containerNumbers != null) {
			Util.replaceDocxText(doc, Constants.CONTAINER_NUMBER_PLACEHOLDER, containerNumbers);
		} else {
			error += "ERROR: Container Numbers not found.\n" + 
					"错误： 找不到Container Numbers.\n";
		}
		
		// Bill of Lading Number
		String billOfLading = obl.getBillOfLadingNumber();
		if (billOfLading != null) {
			Util.replaceDocxText(doc, Constants.BILL_OF_LADING_PLACEHOLDER, billOfLading);
		} else {
			error += "ERROR: Bill of Lading Number not found.\n" + 
					"错误： 找不到Bill of Lading Number.\n";
		}
		
		
		// PI
		ProformaInvoiceModway pi = new ProformaInvoiceModway(pi_path);
		
		String poNum = pi.getPoNumber();
		if (poNum != null) {
			Util.replaceDocxText(doc, Constants.PO_NUMBER_PLACEHOLDER, poNum);
			Util.replaceDocxText(doc, Constants.ENTRY_NUMBER_PLACEHOLDER, poNum);
			
			output_address += "LACEY_ACT-官版  -" + poNum;
	
		} else {
			error += "ERROR: PO # not found.\n" + 
					"错误： 找不到PO #.\n";
		}
		
		
		XWPFTable table = Util.findTableGivenText(doc.getTables(), "9401619000");
		
		List<Item> items = pi.getItems();
		if (items != null && !items.isEmpty()) {
			int lastRowNum = 2;
			for (Item item : items) {
				XWPFTableRow oldRow = table.getRow(lastRowNum);
				CTRow ctRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
				XWPFTableRow newRow = new XWPFTableRow(ctRow, table);
				
				int i = 0;
				for (XWPFTableCell cell : newRow.getTableCells()) {
					for (XWPFParagraph paragraph : cell.getParagraphs()) {
						for (XWPFRun run : paragraph.getRuns()) {
							if (i == 1) {
								run.setText(Double.toString(item.getTotalAmount()), 0);	
							}
							if (i == 2) {
								run.setText(item.getPartNum(), 0);	
							}
							if (i == 7) {
								double qty = Math.round((item.getQuantity() * 0.02) * 100.0) / 100.0;
								run.setText(Double.toString(qty), 0);	
							}

							i++;
						}
					}
				}
				
				table.addRow(newRow, lastRowNum + 1);
				lastRowNum++;
			}
			table.removeRow(2);
		} else {
			error += "ERROR: Product Items not found in given PI.\n" + 
					"错误： 找不到任何产品在PI里.\n";
		}
		
		
		
		if (error.isEmpty()) {
			error = "Generater without error.";
		}
		
		
		output_address = Util.correctFileFormat(".docx", output_address);
		doc.write(new FileOutputStream(output_address));
		doc.close();
		
		return this.error;
	}
}













