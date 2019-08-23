package com.rzheng.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.rzheng.modway.Item;

import io.github.jonathanlink.PDFLayoutTextStripper;
import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;

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
			line.contains(Constants.CONTAINER_SIZE) ||
			line.contains(Constants.PRINTED_ON));
			
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
	
	public static String correctFileFormat(String fileFormat, String filePath) {
		if (filePath != null && !filePath.isEmpty()) {
			if (filePath.substring(filePath.length()-fileFormat.length()).equalsIgnoreCase(fileFormat)) {
				return filePath;
			} else {
				return filePath + fileFormat;
			}
		}	
		return null;
	}

	public static double extractNumberFromAmount(String amount) {
		amount = amount.replaceAll("\\$", "").trim();
		amount = amount.replaceAll(",", "").trim();
		return Double.parseDouble(amount);
	}
	
	public static String readPDF(String pdf_path) {
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
        }
		return text; 
	}
	
	public static String readDocument(String path) {
		if (!path.isEmpty()) {
			if (path.contains(".doc")) {
				if (path.contains(".docx")) {
					return readDocx(path);
				}
				return readDoc(path);
			}
		}
		return null;
	}
	
	private static String readDoc(String doc_path) {
		String text = null;
		try {
			File file = new File(doc_path);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());

			HWPFDocument doc = new HWPFDocument(fis);

			WordExtractor we = new WordExtractor(doc);
//			String[] paragraphs = we.getParagraphText();
			
			text = we.getText();
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return text; 
	}
	
	private static String readDocx(String docx_path) {
		String text = null;
		try {
			File file = new File(docx_path);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());

			XWPFDocument document = new XWPFDocument(fis);

			List<XWPFParagraph> paragraphs = document.getParagraphs();
			
			text = "";
			for (XWPFParagraph para : paragraphs) {
				text += para.getText() + "\n";
			}
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return text; 
	}
	
	// ** Inaccurate reading from low resolution images. e.g. 8 reads $ **
	// Download language packs: https://github.com/tesseract-ocr/tessdata
	public static String readImage(String tessdata_path, String language, String img_path) {
		String text = null;

		ITesseract instance = new Tesseract();
		
		instance.setDatapath(tessdata_path); // "Y:\\Users\\Richard\\spring-tool-suite-4-workspace\\MiaoExcel\\tessdata"
		instance.setLanguage(language); // "eng"

		try {
			text = instance.doOCR(new File(img_path));
		} catch (TesseractException e) {
			e.getMessage();
		}

		return text;
	}
	
	public static String[] fetchModel(String product_file_path, String itemNumber) throws IOException, InvalidFormatException {

		// for any Excel version both .xls and .xlsx
		Workbook wb = WorkbookFactory.create(new File(product_file_path));
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
	
	数�?     	净�?			毛�?			体积
	24.00	1513.20		1644.00 	49.62 
	0.00 	0.00 		0.00 		0.00 
	24.00 	780.00 		900.00 		21.70 
			0.00 		0.00 		0.00 
	48.00 	2293.20 	2544.00 	71.32 

	48	CARTONS	2544	KGS	71.32 CBM

	 */
	
	public static double[] fetchDimensions(String dimension_file_path, String modelNumber, String type, int quantity) throws IOException {

		// for any Excel version both .xls and .xlsx
		Workbook wb = WorkbookFactory.create(new File(dimension_file_path));
		if (wb == null)
			return null;
		
		boolean found = false;

		for (Sheet sheet : wb) {
			
			if (sheet.getSheetName().equals("目录"))
				continue;

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

					System.out.println(row.getRowNum());
					if (row.getCell(0) != null) {
						System.out.println(row.getCell(0).getCellType());
						
						if (row.getCell(0).getCellType() == CellType.STRING)
							System.out.println(row.getCell(0).getRichStringCellValue().getString());
					}
					else {
						System.out.println("NULL cell");
					}
					
					if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.STRING) {
						// End of item / beginning of next item
						if (row.getCell(0).getRichStringCellValue().getString().trim().substring(0, 2).equalsIgnoreCase("YB")) {
							return null;
						}
						
						if (row.getCell(0).getRichStringCellValue().getString().trim().equals(type)) {
							int quantityCol = 7;
							int netWeightCol = 8;
							int grossWeightCol = 9;
							int volumeCol = 10;
							
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
	
	
	public static void copyRow(Workbook workbook, Sheet worksheet, int sourceRowNum, int destinationRowNum) throws EncryptedDocumentException, IOException {

    	// Get the source / new row
        Row newRow = worksheet.getRow(destinationRowNum);
        Row sourceRow = worksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(destinationRowNum);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            ;
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }

        // If there are are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                        )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }
	
	public static Item findItem(List<Item> items, String styleNum, String vendorStyleNum) {
		
		if (items != null && !items.isEmpty() && styleNum != null && vendorStyleNum != null) {
			for (Item item : items) {
				// Style # (Part No.)
				// Vendor Style # (Item #)
				if (item.getPartNum().equalsIgnoreCase(styleNum) && item.getItemNum().equalsIgnoreCase(vendorStyleNum)) {
					return item;
				}
			}
		}
		return null;
	}
	
	public static void replaceDocxText(XWPFDocument doc, String findText, String replaceText) throws IOException {
		replaceDocxTextInParagraphs(doc.getParagraphs(), findText, replaceText);
		replaceDocxTextInTables(doc.getTables(), findText, replaceText);
	}
	
	private static void replaceDocxTextInParagraphs(List<XWPFParagraph> paragraphs, String findText, String replaceText) throws IOException {

		for (XWPFParagraph p : paragraphs) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(findText)) {
		                text = text.replace(findText, replaceText);
		                r.setText(text, 0);
		            }
		        }
		    }
		}
	}
	
	private static void replaceDocxTextInTables(List<XWPFTable> tables, String findText, String replaceText) throws IOException {

		for (XWPFTable tbl : tables) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					// tables within table
					if (cell.getTables() != null) {
						replaceDocxTextInTables(cell.getTables(), findText, replaceText);
					}
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							String text = r.getText(0);
							if (text != null && text.contains(findText)) {
								text = text.replace(findText, replaceText);
								r.setText(text, 0);
							}
						}
					}
				}
			}
		}
	}
	
	public static XWPFTable findTableGivenText(List<XWPFTable> tables, String textInTable) {
		for (XWPFTable tbl : tables) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun r : p.getRuns()) {
							String text = r.getText(0);
							if (text != null && text.equalsIgnoreCase(textInTable)) {
								return tbl;
							}
						}
					}
				}
			}
		}
		return null;
	}
	
	public static String translateCountry(String country) {
		if (country.equalsIgnoreCase(Constants.SAUDI_ARABIA))
			return "沙特阿拉伯";
		if (country.equalsIgnoreCase(Constants.SEOUL))
			return "韩国";
		if (country.equalsIgnoreCase(Constants.US) || country.equalsIgnoreCase(Constants.USA))
			return "美国";
		return country;
	}
	
}
























