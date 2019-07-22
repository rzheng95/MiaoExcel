package com.rzheng.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

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
	
	public static String read(String pdf_path) {
		String text = null;
        try {
            PDFParser pdfParser = new PDFParser(new RandomAccessFile(new File(pdf_path), "r"));
            pdfParser.parse();
            PDDocument pdDocument = new PDDocument(pdfParser.getDocument());
            
            PDFTextStripper pdfTextStripper = new PDFLayoutTextStripper();
            text = pdfTextStripper.getText(pdDocument);
//            String lines[] = text.split("\\r?\\n");
			
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        };
		return text; 
	}
}
