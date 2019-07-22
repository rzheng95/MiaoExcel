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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;

public class ShippingOrder 
{
	private final int CONSIGNEE_ROW = 8, CONSIGNEE_COL = 0;

	private final int NOTIFY_ROW = 15, NOTIFY_COL = 0;

	private final int PORT_OF_DISCHARGE_ROW = 23, PORT_OF_DISCHARGE_COL = 0;

	private final int SEA_AIR_ROW = 11, SEA_AIR_COL = 4;

	private final int PORT_OF_LOADING_ROW = 21, PORT_OF_LOADING_COL = 2;

	private final int DESTINATION_ROW = 23, DESTINATION_COL = 2;

	private final int SHIP_TO_ADDRESS_ROW = 35, SHIP_TO_ADDRESS_COL = 1;

	private final int BILL_OF_LADING_REQUIREMENT_ROW = 14, BILL_OF_LADING_REQUIREMENT_COL = 5;

	private final int PO_ROW = 28, PO_COL = 2;

	private final int CPO_ROW = 29, CPO_COL = 2;

	private final int FORWARDER_ROW = 42, FORWARDER_COL = 1;

	private final int x20_ROW = 16, x20_COL = 3;

	private final int x40_ROW = 17, x40_COL = 3;

	private final int x40HC_ROW = 18, x40HC_COL = 3;
	
	private final int QUANTITY_ROW = 25, QUANTITY_COL = 2;
	
	 //Read the spreadsheet that needs to be updated
	 FileInputStream fsIP= new FileInputStream(new File("Shipping  Order Template.xls"));  
	 //Access the workbook                  
	 HSSFWorkbook wb = new HSSFWorkbook(fsIP);
	 //Access the worksheet, so that we can update / modify it. 
	 HSSFSheet worksheet = wb.getSheetAt(0); 
	 
	 // declare a Cell object
	 Cell cell = null;
	 
	 public ShippingOrder() throws IOException, InvalidFormatException
	 {

	 }
	
	 public String run(String si_pdf_path, String pi_pdf_path, String so_xls_path) throws IOException, InvalidFormatException 
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
	                	if (lines[i].toUpperCase().contains(Constants.PO) && lines[i].toUpperCase().contains(Constants.CO) && lines[i].toUpperCase().contains(Constants.CPO))
	                	{
	                		// ex: PO # 052059 CO # : 1001314 CPO # : DOM#65347
	                		int co_index = lines[i].toUpperCase().indexOf(Constants.CO);
	                		int cpo_inedx = lines[i].toUpperCase().indexOf(Constants.CPO);
	                		String po = lines[i].toUpperCase().substring(0, co_index).trim();
	                		String cpo = lines[i].toUpperCase().substring(cpo_inedx).trim();
	                		
	                		

	                		cell = worksheet.getRow(PO_ROW).getCell(PO_COL);
		           		 	cell.setCellValue(po);
		           		 	
		           		 	cell = worksheet.getRow(CPO_ROW).getCell(CPO_COL);
		           		 	cell.setCellValue(cpo);
		           		 	
		           		 	if(so_xls_path.isEmpty())
		           		 	{
		           		 		String[] arr = po.split(" ");
		           		 		so_xls_path = "Shipping Order " + arr[arr.length-1].trim() + ".xls";
		           		 	}
		           		 	
		           		 	
	                	}
	                	else if (lines[i].toUpperCase().contains(Constants.CONTAINER_SIZE))
	                	{             		
	                		String size = lines[i].toUpperCase().substring(lines[i].toUpperCase().indexOf(Constants.CONTAINER_SIZE) + Constants.CONTAINER_SIZE.length()).trim();

	                		if (size.equalsIgnoreCase(Constants._20)) {
	                			cell = worksheet.getRow(x20_ROW).getCell(x20_COL);
			           		 	cell.setCellValue(1);
	                		} else if (size.equalsIgnoreCase(Constants._40)) {
	                			cell = worksheet.getRow(x40_ROW).getCell(x40_COL);
			           		 	cell.setCellValue(1);
	                		} else if (size.equalsIgnoreCase(Constants._40HC)) {
	                			cell = worksheet.getRow(x40HC_ROW).getCell(x40HC_COL);
			           		 	cell.setCellValue(1);
	                		} 
	                		
	                	}
	                	else if (lines[i].toUpperCase().contains(Constants.CONSIGNEE))
	                	{
	                		String str = lines[i].substring(Constants.CONSIGNEE.length()).trim();
		            		// Access the second cell in second row to update the value
		            		cell = worksheet.getRow(CONSIGNEE_ROW).getCell(CONSIGNEE_COL);
		            		// Get current cell value value and overwrite the value
		           		 	
		           		 	i++;
	                		while (!Util.checkEndLine(lines[i].toUpperCase()))
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
	                	else if (lines[i].toUpperCase().contains(Constants.NOTIFY))
	                	{
	                		if (lines[i].toUpperCase().contains(Constants.ALSO_NOTIFY)) {
	                			i++;
	                			continue;
	                		}
	                		if (lines[i].toUpperCase().contains(Constants._2ND_NOTIFY)) {
	                			i++;
	                			continue;
	                		}
	                		
	                		String str = lines[i].substring(Constants.NOTIFY.length()).trim();
	                		cell = worksheet.getRow(NOTIFY_ROW).getCell(NOTIFY_COL);
	                		
		           		 	i++;
	                		while (!Util.checkEndLine(lines[i].toUpperCase()))
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
	                	else if (lines[i].toUpperCase().contains(Constants.PORT_OF_LOADING))
	                	{
	                		cell = worksheet.getRow(PORT_OF_LOADING_ROW).getCell(PORT_OF_LOADING_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(Constants.PORT_OF_LOADING.length()).trim());
	                	}
	                	else if (lines[i].toUpperCase().contains(Constants.PORT_OF_DISCHARGE))
	                	{
	                		cell = worksheet.getRow(PORT_OF_DISCHARGE_ROW).getCell(PORT_OF_DISCHARGE_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(Constants.PORT_OF_DISCHARGE.length()).trim());
		           		 	
		           		 	cell = worksheet.getRow(SEA_AIR_ROW).getCell(SEA_AIR_COL);
		           		 	if (lines[i].toUpperCase().substring(Constants.PORT_OF_DISCHARGE.length()-1).contains(Constants.SEA_AIR)) 
			           		 	cell.setCellValue(Constants.SEA_AIR);
		           		 	else
		           		 	cell.setCellValue("AIR");
	                	}
	                	else if (lines[i].toUpperCase().contains(Constants.DESTINATION))
	                	{
	                		cell = worksheet.getRow(DESTINATION_ROW).getCell(DESTINATION_COL);
		           		 	cell.setCellValue(lines[i].toUpperCase().substring(Constants.DESTINATION.length()).trim());
	                	}
	                	else if (lines[i].toUpperCase().contains(Constants.SHIP_TO_ADDRESS))
	                	{
	                		String str = lines[i].substring(Constants.SHIP_TO_ADDRESS.length()).trim();
	                		cell = worksheet.getRow(SHIP_TO_ADDRESS_ROW).getCell(SHIP_TO_ADDRESS_COL);
//                			System.out.println(lines[i]);

		           		 	i++;
	                		while (!Util.checkEndLine(lines[i].toUpperCase()))
	                		{
//	                			System.out.println(lines[i]);
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
	                	else if (lines[i].toUpperCase().contains(Constants.FORWARDER))
	                	{
	                		String str = lines[i].substring(Constants.FORWARDER.length()).trim();
	                		cell = worksheet.getRow(FORWARDER_ROW).getCell(FORWARDER_COL);

		           		 	i++;
	                		while (!Util.checkEndLine(lines[i].toUpperCase()))
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
	                	else if (lines[i].toUpperCase().contains(Constants.ISSUE) && lines[i].toUpperCase().contains(Constants.BL))
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
	            
	            document.close();
	        }
		 

		 // PI
		 String pi = Util.read(pi_pdf_path);
		 
		 String[] lines = pi.split("\\r?\\n");
		 int i = 0;
		 
		 // remove extra spaces
		 while ( i < lines.length ) {
			 lines[i] = lines[i].trim().replaceAll(" +", " ");
			 i++;
		 }
		 
		 i = 0;
		 
		 while ( i < lines.length ) {
			
			 // U3446-20- Silver Sofa KD / KD sofa argent 9401.61 67.80 24 $314.00 $7,536.00
			 if (Util.countString(lines[i], "$", 2) && Util.countString(lines[i], "-", 2))
			 {
				 
				 String item = lines[i];
				 String[] arr = item.split(" ");
				 
				 String itemNum = arr[0];
				 int quantity = Integer.parseInt(arr[arr.length-3]);
				 
				 
				 i++;
				 arr = lines[i].split(" ");
				 itemNum += arr[0];
//				 System.out.println(itemNum);
				 
				 String[] modelNum = Util.fetchModel(itemNum);

				 if(modelNum != null) {
					 
					 Util.fetchStats(modelNum[0].substring(0, 6), modelNum[1], quantity);

				 } else {
					 return "ERROR: Can not find item number " + itemNum + "\n错误： 找不到客户型号" + itemNum;
				 }
				 
			 }
			 if (lines[i].contains(Constants.TOTAL)) {
				 
					if (lines[i].contains(Constants.TOTAL_EXCL_TAX)) {
	           		 	i++;
	           		 	continue;
	           		 	
					} else if (lines[i].contains(Constants.SUB_TOTAL)) {
						i++;
						continue;
					}
					
					// TOTAL 2388.96 48
					String[] arr = lines[i].split(" ");
					
					if(arr.length == 3) {
						cell = worksheet.getRow(QUANTITY_ROW).getCell(QUANTITY_COL);
						cell.setCellValue(Integer.parseInt(arr[arr.length-1]));
					}

				}
			 
			 
			 
			 
			 
			 
			 
			 
			 
			 
			 
			 
			 i++;
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
		 
		 return "Success!";
	}
	
}
