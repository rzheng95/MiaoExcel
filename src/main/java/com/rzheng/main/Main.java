package com.rzheng.main;


import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFrame;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.rzheng.excel.CustomsDeclaration;
import com.rzheng.excel.ShippingOrder;

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

//		CustomsDeclaration cd = new CustomsDeclaration("052059 SI.pdf", "052059 PI.pdf", "", "Customs Declaration Template.xls", "INYB2019US0449");
//		String cd_error = cd.run();
//		System.out.println(cd_error);
		
		JFrame frame = new JFrame("My First GUI");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(300,300);
		JButton button = new JButton("Press");
		frame.getContentPane().add(button); // Adds Button to content pane of frame
		frame.setVisible(true);
	}
}