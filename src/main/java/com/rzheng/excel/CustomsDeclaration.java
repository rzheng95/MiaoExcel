package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Scanner;

import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import io.github.jonathanlink.PDFLayoutTextStripper;

public class CustomsDeclaration 
{
	private final String TOTAL = "TOTAL";
	
	private final String SUB_TOTAL = "SUB TOTAL";
	
	private final String TOTAL_EXCL_TAX = "TOTAL EXCL. TAX";
	private final int TOTAL_EXCL_TAX_ROW = 15, TOTAL_EXCL_TAX_COL = 6;
	
	// Read the spreadsheet that needs to be updated
	FileInputStream fsIP = new FileInputStream(new File("Customs Declaration Template.xls"));
	// Access the workbook
	HSSFWorkbook wb = new HSSFWorkbook(fsIP);
	// Access the worksheet, so that we can update / modify it.
	HSSFSheet worksheet = wb.getSheetAt(0);

	// declare a Cell object
	Cell cell = null;

	public CustomsDeclaration(String si_pdf_path, String pi_pdf_path, String cd_xls_path) throws IOException
	{
		this.contract(pi_pdf_path, cd_xls_path);
	}
	
	public void contract(String pi_pdf_path, String cd_xls_path) throws IOException
	{
		try (PDDocument document = PDDocument.load(new File(pi_pdf_path))) {
			document.getClass();

			if (!document.isEncrypted()) {

				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);

				PDFTextStripper tStripper = new PDFTextStripper();

				String pdfFileInText = tStripper.getText(document);
//	                System.out.println("Text:" + pdfFileInText);

				// split by whitespace
				String lines[] = pdfFileInText.split("\\r?\\n");
				for (int i = 0; i < lines.length; i++)
					lines[i] = lines[i].toUpperCase();
				int i = 0;

				while (i < lines.length) {
					if (lines[i].contains(TOTAL)) {
						if (lines[i].contains(TOTAL_EXCL_TAX)) {
		            		cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
		            		double amount = extractNumberFromString(lines[i].substring(TOTAL_EXCL_TAX.length()));
		           		 	cell.setCellValue(amount);
//		            		cell.setCellType(CellType.NUMERIC);
//		            		CellStyle cs = wb.createCellStyle();
//		            		cs.setDataFormat((short)7);
//		            		cell.setCellStyle(cs);
							
		           		 	// set total amount
		           		 	cell = worksheet.getRow(TOTAL_EXCL_TAX_ROW).getCell(TOTAL_EXCL_TAX_COL);
		           		 	cell.setCellValue(amount);
		           		 	
		           		 	
						} else if (lines[i].contains(SUB_TOTAL)) {
							i++;
							continue;
						}
						System.out.println(lines[i]);
					}

					i++;
				}

			}

		}
		HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
		// Close the InputStream
		fsIP.close();

		if (!cd_xls_path.contains(".xls") && !cd_xls_path.isEmpty())
			cd_xls_path = cd_xls_path + ".xls";
		if (cd_xls_path.contains(".xlsx"))
			cd_xls_path = cd_xls_path.substring(0, cd_xls_path.length() - 1);

		// Open FileOutputStream to write updates
		FileOutputStream output_file = new FileOutputStream(new File(cd_xls_path));
		// write changes
		wb.write(output_file);
		// close the stream
		output_file.close();
	}
	
	private double extractNumberFromString(String amount) {
		amount = amount.replaceAll("\\$", "").trim();
		amount = amount.replaceAll(",", "").trim();
		return Double.parseDouble(amount);
	}
	
	private boolean countDollarSign(String line, int countThreshold) {
		int count = 0;
		StringBuilder sb = new StringBuilder(line);
		for (int i = 0; i < countThreshold; i++) {
			if(sb.indexOf("$") != -1) {
				sb = sb.deleteCharAt(line.indexOf('$'));
				count++;
			}
			if(count >= countThreshold)
				return true;
			
		}
		return false;	
	}
}































