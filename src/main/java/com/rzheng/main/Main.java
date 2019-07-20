package com.rzheng.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.rzheng.excel.ShippingOrder;

//import com.aspose.pdf.Document;
//import com.aspose.pdf.ExcelSaveOptions;;

public class Main 
{
	public static void main(String[] args) throws IOException 
	{	
		ShippingOrder so = new ShippingOrder("SI.pdf", "", "");
	}
}