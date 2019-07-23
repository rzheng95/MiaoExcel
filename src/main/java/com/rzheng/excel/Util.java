package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.github.jonathanlink.PDFLayoutTextStripper;

public final class Util {
	
	public static boolean checkEndLine(String line) {
		return ( line.contains(Constants.CONSIGNEE) ||
			line.contains(Constants.NOTIFY) ||
			line.contains(Constants.ALSO_NOTIFY) ||
			line.contains(Constants._2ND_NOTIFY) ||
			line.contains(Constants.PORT_OF_DISCHARGE) ||
			line.contains(Constants.PORT_OF_LOADING) ||
			line.contains(Constants.DESTINATION) ||
			line.contains(Constants.SHIP_TO_ADDRESS) ||
			line.contains(Constants.SELECTION_CRITERIA) ||
			line.contains(Constants.FORWARDER) ||
			line.contains(Constants.CARRIER) ||
			line.contains(Constants.CONTAINER_SIZE));
			
	}
	
	public static boolean countString(String line, String string, int countThreshold) {

		int count = 0;
		StringBuilder sb = new StringBuilder(line);
		for (int i = 0; i < countThreshold; i++) {
			if(sb.indexOf(string) != -1) {
				sb = sb.deleteCharAt(line.indexOf(string));
				count++;
			}
			if(count >= countThreshold)
				return true;
		}
		return false;	
	}
	
	public static String correctXlsFilename(String xls_filename) {
		if (!xls_filename.contains(".xls") && !xls_filename.isEmpty())
			xls_filename = xls_filename + ".xls";
		if (xls_filename.contains(".xlsx"))
			xls_filename = xls_filename.substring(0, xls_filename.length() - 1);
		return xls_filename;
	}
	
	public static double extractNumberFromAmount(String amount) {
		amount = amount.replaceAll("\\$", "").trim();
		amount = amount.replaceAll(",", "").trim();
		return Double.parseDouble(amount);
	}
	
	public static String read(String pdf_path) {
		String text = null;
        try {
            PDFParser pdfParser = new PDFParser(new RandomAccessFile(new File(pdf_path), "r"));
            pdfParser.parse();
            PDDocument pdDocument = new PDDocument(pdfParser.getDocument());
            
            PDFTextStripper pdfTextStripper = new PDFLayoutTextStripper();
            text = pdfTextStripper.getText(pdDocument);
            
            
            pdDocument.close();
			
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        };
		return text; 
	}
	
	public static String[] fetchModel(String itemNumber) throws IOException, InvalidFormatException {

		// for any Excel version both .xls and .xlsx
		Workbook wb = WorkbookFactory.create(new File("magnussen 产品对照表 201905025.xlsx"));
		Sheet worksheet = wb.getSheetAt(0);

		for (Row row : worksheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(itemNumber)) {
						
						wb.close();
						return new String[] {
								worksheet.getRow(row.getRowNum()).getCell(0).toString(), 
								worksheet.getRow(row.getRowNum()).getCell(1).toString()
						};
					}
				}
			}
		}

		wb.close();
		return null;
	}
	
	/*
	
	数量     	净重			毛重			体积
	24.00	1513.20		1644.00 	49.62 
	0.00 	0.00 		0.00 		0.00 
	24.00 	780.00 		900.00 		21.70 
			0.00 		0.00 		0.00 
	48.00 	2293.20 	2544.00 	71.32 

	48	CARTONS	2544	KGS	71.32 CBM

	 */
	
	public static double[] fetchStats(String modelNumber, String type, int quantity) throws IOException {
//		System.out.println(modelNumber);
//		System.out.println(type);
//		System.out.println(quantity);
		// for any Excel version both .xls and .xlsx
		Workbook wb = WorkbookFactory.create(new File("净毛体统计2016.09.07.xls"));
		
		boolean found = false;

		for (Sheet sheet : wb) {

			for (Row row : sheet) {
				
				if (!found) {
					for (Cell cell : row) {

						if (cell.getCellType() == CellType.STRING) {
							if (cell.getRichStringCellValue().getString().trim().equals(modelNumber)) {
								found = true;
								break;
							}
						}

					}
				}
				// found
				else {

					int quantityCol = 7;
					int netWeightCol = 8;
					int grossWeightCol = 9;
					int volumeCol = 10;
					// NEED TO CHECK WHEN TO STOP
					// NEED TO CHECK WHEN TO STOP
					// NEED TO CHECK WHEN TO STOP
					if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.STRING) {
						if (row.getCell(0).getRichStringCellValue().getString().trim().equals(type)) {
							
							row.getCell(quantityCol).setCellValue(quantity);
							
							HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
							XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
							
							double netWeight =  Math.round(row.getCell(netWeightCol).getNumericCellValue() * 100.0) / 100.0;
							double grossWeight =  Math.round(row.getCell(grossWeightCol).getNumericCellValue() * 100.0) / 100.0;
							double volume =  Math.round(row.getCell(volumeCol).getNumericCellValue() * 100.0) / 100.0;
							return new double[] {netWeight, grossWeight, volume};
						}
					}

				}
			}

		}

		wb.close();
		return null;
	}
	
}
























