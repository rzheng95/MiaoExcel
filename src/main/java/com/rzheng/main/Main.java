package com.rzheng.main;


import java.io.IOException;

import javax.swing.SwingUtilities;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;

import com.rzheng.gui.CustomsClearanceGUI;
import com.rzheng.gui.CustomsClearanceModwayGUI;
import com.rzheng.gui.CustomsDeclarationGUI;
import com.rzheng.gui.LaceyActAmendmentGUI;
import com.rzheng.gui.ShippingOrderGUI;
import com.rzheng.magnussen.CustomsClearance;
import com.rzheng.modway.CustomsClearanceModway;
import com.rzheng.modway.LaceyActAmendment;
import com.rzheng.util.Util;

public class Main
{
	public static void main(String[] args) throws InvalidFormatException, IOException, XmlException 
	{	
		// 051490 X
		// 052059
		// 051488 X
		// 051487 X
		// 051338 
		// 051336
//		String so_error = new ShippingOrder().run("052059 SI.pdf", "052059 PI.pdf", "", "Shipping Order Template.xls");
//		System.out.println(so_error);

//		CustomsClearance cc = new CustomsClearance("magnussen 产品对照表 201905025.xlsx", "净毛体统计2016.09.07(1).xls", "052059 SI.pdf", "052059 PI.pdf", 
//				"Unitex Logistics", "指代/Unitex Logistics Ltd 051490.docx", "cc_test", "Customs Clearance Template.xlsx", "invoice number", "container number", "seal number");
//		String cc_error = cc.run();
//		System.out.println(cc_error);
		
//		CustomsClearanceModway cc = new CustomsClearanceModway("modway/9395/0009395-PI-MODWAY-041919(1).xls", 
//				"modway/9395/0009395 HES19050515-海运提单(代理).PDF", "modway/9395/9395分货-有净毛体(1).xls", "modway/Modway Customs Clearance Template.xls", "modway/cc_modway_test", "", "etd", "eta");
//		
//		System.out.println(cc.run());
		
//		LaceyActAmendment laa = new LaceyActAmendment("modway/9395/0009395-PI-MODWAY-041919(1).xls", "modway/9395/0009395 HES19050515-海运提单(代理).PDF", "modway/Lacey Act Template.docx", "modway/", "TEMP ETA !!1");
//		System.out.println(laa.run());
		
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
//				new Login();
//				new ShippingOrderGUI();
//				new CustomsDeclarationGUI();
//				new CustomsClearanceGUI();
//				new CustomsClearanceModwayGUI();
				new LaceyActAmendmentGUI();
			}
			
		});

	}
}





















