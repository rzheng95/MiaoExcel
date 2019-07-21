package com.rzheng.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class ShippingOrder 
{
	private final String CONSIGNEE = "CONSIGNEE:";
	private final int CONSIGNEE_ROW = 8, CONSIGNEE_COL = 0;
	
	private final String NOTIFY = "NOTIFY: ";
	private int NOTIFY_ROW = 15, NOTIFY_COL = 0;
	private final String ALSO_NOTIFY = "ALSO NOTIFY:";
	
	private final String PORT_OF_DISCHARGE = "PORT OF DISCHARGE:";
	private final int PORT_OF_DISCHARGE_ROW = 23, PORT_OF_DISCHARGE_COL = 0;
	
	private final String SEA_AIR = "SEA";
	private final int SEA_AIR_ROW = 11, SEA_AIR_COL = 4;
		
	private final String PORT_OF_LOADING = "PORT OF LOADING:";
	private final int PORT_OF_LOADING_ROW = 21, PORT_OF_LOADING_COL = 2;
	
	private final String DESTINATION = "DESTINATION:";
	private final int DESTINATION_ROW = 23, DESTINATION_COL = 2;
	
	private final String SHIP_TO_ADDRESS = "SHIP-TO ADDRESS:";
	private int SHIP_TO_ADDRESS_ROW = 35, SHIP_TO_ADDRESS_COL = 1;
	private final String SELECTION_CRITERIA = "SELECTION CRITERIA:";
	
	private final int BILL_OF_LADING_REQUIREMENT_ROW = 14, BILL_OF_LADING_REQUIREMENT_COL = 5;
	private final String ISSUE = "ISSUE";
	private final String BL = "B/L";
	
	private final String QUANTITY = "";
	
	private final String PO = "PO #";
	private final int PO_ROW = 28, PO_COL = 2;
	private final String CPO = "CPO #";
	private final int CPO_ROW = 29, CPO_COL = 2;
	
	private final String FORWARDER = "FORWARDER:";
	private int FORWARDER_ROW = 42, FORWARDER_COL = 1;
	private final String CARRIER = "CARRIER:";
	
	private final String CONTAINER_SIZE = "CONTAINER SIZE:";
	private final String _20 = "20";
	private final int x20_ROW = 16, x20_COL = 3;
	private final String _40 = "40";
	private final int x40_ROW = 17, x40_COL = 3;
	private final String _40HC = "40HC";
	private final int x40HC_ROW = 18, x40HC_COL = 3;
	
	 //Read the spreadsheet that needs to be updated
	 FileInputStream fsIP= new FileInputStream(new File("Shipping  Order Template.xls"));  
	 //Access the workbook                  
	 HSSFWorkbook wb = new HSSFWorkbook(fsIP);
	 //Access the worksheet, so that we can update / modify it. 
	 HSSFSheet worksheet = wb.getSheetAt(0); 
	 
	 // declare a Cell object
	 Cell cell = null;
	 
	 public ShippingOrder(String si_pdf_path, String pi_pdf_path, String so_xls_path) throws IOException
	 {
		 this.run(si_pdf_path, pi_pdf_path, so_xls_path);
	 }
	
	 public void run(String si_pdf_path, String pi_pdf_path, String so_xls_path) throws IOException 
	 {

		 try (PDDocument document = PDDocument.load(new File(si_pdf_path))) 
		 {
	            document.getClass();
	
	            if (!document.isEncrypted()) 
	            {
				
	                PDFTextStripperByArea stripper = new PDFTextStripperByArea();
	                stripper.setSortByPosition(true);
	
	                PDFTextStripper tStripper = new PDFTextStripper();
	
	                String pdfFileInText = tStripper.getText(document);
//	                System.out.println("Text:" + pdfFileInText);
	                
					// split by whitespace
	                String lines[] = pdfFileInText.split("\\r?\\n");

	                int i = 0;
	                
	                while ( i < lines.length )
	                {
	                	if (lines[i].toUpperCase().contains(PO) && lines[i].toUpperCase().contains(CPO))
	                	{
	                		String[] arr = lines[i].toUpperCase().split(" ");
	                		cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
		           		 	cell.setCellValue(arr[0] + arr[1] + " " + arr[2]);
		           		 	if(so_xls_path.isEmpty())
		           		 		so_xls_path = "Shipping Order " + arr[2] + ".xls";
		           		 		
		           		 	cell = worksheet.getRow(CPO_ROW).getCell(CPO_COL);
		           		 	cell.setCellValue(arr[6] + arr[7] + " " + arr[8]);
	                	}
	                	else if (lines[i].toUpperCase().contains(CONTAINER_SIZE))
	                	{             		
	                		String size = lines[i].toUpperCase().substring(lines[i].toUpperCase().indexOf(CONTAINER_SIZE) + CONTAINER_SIZE.length()).trim();

	                		if (size.equalsIgnoreCase(_20)) {
	                			cell = worksheet.getRow(x20_ROW).getCell(x20_COL);
			           		 	cell.setCellValue(1);
	                		} else if (size.equalsIgnoreCase(_40)) {
	                			cell = worksheet.getRow(x40_ROW).getCell(x40_COL);
			           		 	cell.setCellValue(1);
	                		} else if (size.equalsIgnoreCase(_40HC)) {
	                			cell = worksheet.getRow(x40HC_ROW).getCell(x40HC_COL);
			           		 	cell.setCellValue(1);
	                		} 
	                		
	                	}
	                	else if (lines[i].toUpperCase().contains(CONSIGNEE))
	                	{
	                		String str = lines[i].substring(CONSIGNEE.length()).trim();
		            		// Access the second cell in second row to update the value
		            		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
		            		// Get current cell value value and overwrite the value
		           		 	
		           		 	i++;
	                		while (!lines[i].toUpperCase().contains(NOTIFY))
	                		{
	                			if (lines[i].trim().isEmpty())
	                			{
	                				i++;
	                				continue;
	                			}
	                			str += "\n" + lines[i].trim();
			           		 	i++;
	                		}
	                		i--;
	                		cell.setCellValue(str);
	                	}
	                	else if (lines[i].toUpperCase().contains(NOTIFY))
	                	{
	                		String str = lines[i].substring(NOTIFY.length()).trim();
	                		cell = worksheet.getRow(NOTIFY_ROW).getCell(NOTIFY_COL);
		           		 	
		           		 	i++;
	                		while (!lines[i].toUpperCase().contains(ALSO_NOTIFY))
	                		{
	                			if (lines[i].trim().isEmpty())
	                			{
	                				i++;
	                				continue;
	                			}
	                			str += "\n" + lines[i].trim();
			           		 	i++;
	                		}
	                		i--;
	                		cell.setCellValue(str);
	                	}
	                	else if (lines[i].toUpperCase().contains(PORT_OF_LOADING))
	                	{
	                		cell = worksheet.getRow(PORT_OF_LOADING_ROW).getCell(PORT_OF_LOADING_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(PORT_OF_LOADING.length()).trim());
	                	}
	                	else if (lines[i].toUpperCase().contains(PORT_OF_DISCHARGE))
	                	{
	                		cell = worksheet.getRow(PORT_OF_DISCHARGE_ROW).getCell(PORT_OF_DISCHARGE_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(PORT_OF_DISCHARGE.length()).trim());
		           		 	
		           		 	cell = worksheet.getRow(SEA_AIR_ROW).getCell(SEA_AIR_COL);
		           		 	if (lines[i].toUpperCase().substring(PORT_OF_DISCHARGE.length()-1).contains(SEA_AIR)) 
			           		 	cell.setCellValue(SEA_AIR);
		           		 	else
		           		 	cell.setCellValue("AIR");
	                	}
	                	else if (lines[i].toUpperCase().contains(DESTINATION))
	                	{
	                		cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(DESTINATION.length()).trim());
	                	}
	                	else if (lines[i].toUpperCase().contains(SHIP_TO_ADDRESS))
	                	{
	                		String str = lines[i].substring(SHIP_TO_ADDRESS.length()).trim();
	                		cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
		           		 	
		           		 	i++;
	                		while (!lines[i].toUpperCase().contains(SELECTION_CRITERIA))
	                		{
	                			if (lines[i].trim().isEmpty())
	                			{
	                				i++;
	                				continue;
	                			}
	                			str += "\n" + lines[i].trim();
			           		 	i++;
	                		}
	                		i--;
	                		cell.setCellValue(str);
	                	}
	                	else if (lines[i].toUpperCase().contains(FORWARDER))
	                	{
	                		String str = lines[i].substring(FORWARDER.length()).trim();
	                		cell = worksheet.getRow(FORWARDER_ROW).getCell(FORWARDER_COL);

		           		 	i++;
	                		while (!lines[i].toUpperCase().contains(CARRIER))
	                		{
	                			if (lines[i].trim().isEmpty())
	                			{
	                				i++;
	                				continue;
	                			}
	                			str += "\n" + lines[i].trim();
			           		 	i++;
	                		}
	                		i--;
	                		cell.setCellValue(str);
	                	}
	                	else if (lines[i].toUpperCase().contains(ISSUE) && lines[i].toUpperCase().contains(BL))
	                	{
	                		
	                		String bolr = "";
	                		if(lines[i].toUpperCase().contains(","))
	                		{
	                			String[] arr = lines[i].toUpperCase().split(",");
	                			bolr = arr[0].trim();
	                			if(bolr.contains(" ")) {
	                				arr = bolr.split(" ");
	                				if(arr.length > 2)
	                					bolr = arr[1] + " " + arr[2];
	                			}
	                		}
	                		cell = worksheet.getRow(BILL_OF_LADING_REQUIREMENT_ROW).getCell(BILL_OF_LADING_REQUIREMENT_COL);
		           		 	cell.setCellValue(bolr);
	                	}
	                	
	                	
	                	
	                	i++;
	                }
	
	            }
	
	        }
		 
		 //Close the InputStream  
		 fsIP.close(); 
		 
		 if(!so_xls_path.contains(".xls") && !so_xls_path.isEmpty())
			 so_xls_path = so_xls_path + ".xls";
		 if(so_xls_path.contains(".xlsx"))
			 so_xls_path = so_xls_path.substring(0, so_xls_path.length()-1);
		 
		//Open FileOutputStream to write updates
		 FileOutputStream output_file =new FileOutputStream(new File(so_xls_path));  
		 //write changes
		 wb.write(output_file);
		 //close the stream
		 output_file.close();
	}
	
}
