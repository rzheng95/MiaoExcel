package com.rzheng.main;


import java.io.IOException;

import com.rzheng.excel.CustomsDeclaration;
import com.rzheng.excel.ShippingOrder;

public class Main 
{
	public static void main(String[] args) throws IOException 
	{	
		ShippingOrder so = new ShippingOrder("SI.pdf", "", "");
//		CustomsDeclaration cd = new CustomsDeclaration("SI.pdf", "PI.pdf", "test");
	}
}