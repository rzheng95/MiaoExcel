package com.rzheng.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import io.github.jonathanlink.PDFLayoutTextStripper;

public class PDFReader {
	public static void main(String[] args) {
		
		new PDFReader().read("PI.pdf");
	}
	
	public String read(String pdf_path) {
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