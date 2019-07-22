package com.rzheng.main;


import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.rzheng.excel.CustomsDeclaration;
import com.rzheng.excel.ShippingOrder;

public class Main 
{
	public static void main(String[] args) throws InvalidFormatException, IOException 
	{	
		// 051490
		// 052059
		// 051488
		// 051487
		// 051338
		// 051336
		String so_error = new ShippingOrder().run("052059 SI.pdf", "052059 PI.pdf", "");
		System.out.println(so_error);

//		CustomsDeclaration cd = new CustomsDeclaration("052059 SI.pdf", "052059 PI.pdf", "", "INYB2019US0449");
	}
}