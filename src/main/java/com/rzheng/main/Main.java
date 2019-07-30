package com.rzheng.main;


import java.io.IOException;

import javax.swing.SwingUtilities;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.rzheng.excel.CustomsClearance;
import com.rzheng.gui.CustomsDeclarationGUI;
import com.rzheng.gui.ShippingOrderGUI;
import com.rzheng.util.Util;

public class Main
{
	public static void main(String[] args) throws InvalidFormatException, IOException 
	{	
		// 051490 X
		// 052059
		// 051488 X
		// 051487 X
		// 051338 
		// 051336
//		String so_error = new ShippingOrder().run("052059 SI.pdf", "052059 PI.pdf", "", "Shipping Order Template.xls");
//		System.out.println(so_error);

		CustomsClearance cc = new CustomsClearance("magnussen 产品对照表 201905025.xlsx", "净毛体统计2016.09.07(1).xls", "052059 SI.pdf", "052059 PI.pdf", 
				"Unitex Logistics", "指代/Unitex Logistics Ltd 051490.docx", "cc_test", "Customs Clearance Template.xlsx", "invoice number", "container number", "seal number");
		String cc_error = cc.run();
		System.out.println(cc_error);
		
//		String text = Util.readDocument("指代/Unitex Logistics Ltd 051490.docx");
//		System.out.println(text);
		
		
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
//				new Login();
//				new ShippingOrderGUI();
//				new CustomsDeclarationGUI();
			}
			
		});

	}
}





















