package com.rzheng.main;


import java.io.IOException;

import com.rzheng.excel.CustomsDeclaration;
import com.rzheng.excel.ShippingOrder;

public class Main 
{
	public static void main(String[] args) throws IOException 
	{	
//		ShippingOrder so = new ShippingOrder("051488 SI.pdf", "", "");
		CustomsDeclaration cd = new CustomsDeclaration("052059 SI.pdf", "052059 PI.pdf", "", "INYB2019US0449");
	}
}