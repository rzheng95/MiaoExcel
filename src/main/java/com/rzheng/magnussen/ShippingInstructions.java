package com.rzheng.magnussen;

import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import com.rzheng.util.Constants;
import com.rzheng.util.Util;

public class ShippingInstructions {

	public static void main(String[] args) {
		ShippingInstructions si = new ShippingInstructions("052345 SI.pdf");
		System.out.println(si.getBillOfLadingRequirement());
	}

	private String si_pdf_path;

	public ShippingInstructions(String si_pdf_path) {
		this.si_pdf_path = si_pdf_path;

	}

	private String[] readLines() {

		try (PDDocument document = PDDocument.load(new File(this.si_pdf_path))) {
			document.getClass();

			if (!document.isEncrypted()) {

				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);

				PDFTextStripper tStripper = new PDFTextStripper();

				String pdfFileInText = tStripper.getText(document);
//	            System.out.println("Text:" + pdfFileInText);
				
				document.close();
				// split by whitespace
				return pdfFileInText.split("\\r?\\n");
			}
		} catch (InvalidPasswordException e) {
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
	}

	public String getConsignee() {
		
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;
		String consignee = null;
		

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.CONSIGNEE)) {
				consignee = lines[i].substring(Constants.CONSIGNEE.length()).trim();
				// Access the second cell in second row to update the value

				// Get current cell value value and overwrite the value

				i++;
				while (!Util.checkEndLine(lines[i].toUpperCase())) {
					if (lines[i].trim().isEmpty()) {
						i++;
						continue;
					}
					consignee += "\n" + lines[i].trim();
					i++;
				}
				break;
			}

			i++;
		}

		return consignee;
	}

	public String getNotifyParty() {
		
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;
		String notifyParty = null;


		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.NOTIFY)) {
				if (lines[i].toUpperCase().contains(Constants.ALSO_NOTIFY)) {
					i++;
					continue;
				}
				if (lines[i].toUpperCase().contains(Constants._2ND_NOTIFY)) {
					i++;
					continue;
				}

				notifyParty = lines[i].substring(Constants.NOTIFY.length()).trim();

				i++;
				while (!Util.checkEndLine(lines[i].toUpperCase())) {
					if (lines[i].trim().isEmpty()) {
						i++;
						continue;
					}
					notifyParty += "\n" + lines[i].trim();
					i++;
				}
				break;
			}

			i++;
		}

		return notifyParty;
	}

	public String getShipToAddress() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;
		String shipToAddress = null;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.SHIP_TO_ADDRESS)) {
				shipToAddress = lines[i].substring(Constants.SHIP_TO_ADDRESS.length()).trim();

				i++;
				while (!Util.checkEndLine(lines[i].toUpperCase())) {
					if (lines[i].trim().isEmpty()) {
						i++;
						continue;
					}
					shipToAddress += "\n" + lines[i].trim();
					i++;
				}
				break;
			}

			i++;
		}

		return shipToAddress;
	}

	public String getForwarder() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;
		String forwarder = null;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.FORWARDER)) {
				forwarder = lines[i].substring(Constants.FORWARDER.length()).trim();

				i++;
				while (!Util.checkEndLine(lines[i].toUpperCase())) {
					if (lines[i].trim().isEmpty()) {
						i++;
						continue;
					}
					forwarder += "\n" + lines[i].trim();
					i++;
				}
				break;
			}
			i++;
		}

		return forwarder;
	}

	public String getPortOfLoading() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PORT_OF_LOADING)) {
				return lines[i].toUpperCase().substring(Constants.PORT_OF_LOADING.length()).trim();
			}
			i++;
		}
		return null;
	}
	
	public String getPortOfLoadingCountry() {
		String portOfLoading = getPortOfLoading();
		if (portOfLoading != null) {
			if (portOfLoading.contains(",")) {
				String[] arr = portOfLoading.split(",");
				if (arr != null && arr.length == 2) {
					return arr[1].trim();
				}
			}
		}
		
		return null;
	}

	public String getPortOfDischarge() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PORT_OF_DISCHARGE)) {
				return lines[i].toUpperCase().substring(Constants.PORT_OF_DISCHARGE.length()).trim();
			}
			i++;
		}
		return null;
	}

	public String getDestination() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.DESTINATION)) {
				return lines[i].toUpperCase().substring(Constants.DESTINATION.length()).trim();
			}
			i++;
		}

		return null;
	}
	
	public String getPoNumber() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PO) && lines[i].toUpperCase().contains(Constants.CO)
					&& lines[i].toUpperCase().contains(Constants.CPO)) {
				// ex: PO # 052059 CO # : 1001314 CPO # : DOM#65347
				int co_index = lines[i].toUpperCase().indexOf(Constants.CO);
				return lines[i].toUpperCase().substring(0, co_index).trim();
			}
			i++;
		}

		return null;
	}
	
	public String getCpoNumber() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PO) && lines[i].toUpperCase().contains(Constants.CO)
					&& lines[i].toUpperCase().contains(Constants.CPO)) {
				// ex: PO # 052059 CO # : 1001314 CPO # : DOM#65347
				int cpo_inedx = lines[i].toUpperCase().indexOf(Constants.CPO);
				return lines[i].toUpperCase().substring(cpo_inedx).trim();
				
			}
			i++;
		}

		return null;
	}
	
	public String getContainerSize() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.CONTAINER_SIZE)) {
				return lines[i].toUpperCase()
						.substring(lines[i].toUpperCase().indexOf(Constants.CONTAINER_SIZE)
								+ Constants.CONTAINER_SIZE.length())
						.trim();
			}
			i++;
		}

		return null;
	}
	
	public String getBillOfLadingRequirement() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.ISSUE) && !lines[i].toUpperCase().contains(Constants.ISSUED)) {
				System.out.println(lines[i]);
				if (lines[i].contains(",")) {
					String[] arr = lines[i].split(",");
					if (arr != null && arr.length >= 1) {
						int index = arr[0].indexOf(Constants.ISSUE);
						return arr[0].substring(index + Constants.ISSUE.length()).trim();
					}
					
				} else {
					int index = lines[i].indexOf(Constants.ISSUE);
					return lines[i].substring(index + Constants.ISSUE.length()).trim();
				}
				
			}
			i++;
		}

		return null;
	}
	
	public String getCarrier() {
		String[] lines = this.readLines();
		if (lines == null)
			return null;
		
		int i = 0;
		
		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.CARRIER) && !lines[i].toUpperCase().contains(Constants.BACKUP_CARRIER)) {
				return lines[i].substring(Constants.CARRIER.length()).trim();
			}
			
			i++;
		}
		return null;
	}

	public String getDestinationCity() {
		String destination = this.getDestination();
		if(destination != null) {
			String[] arr = null;
			if (destination.contains(",")) {
				arr = destination.split(",");	
			} else if  (destination.contains("-")) {
				arr = destination.split("-");	
			} else if  (destination.contains("–")) {
				arr = destination.split("–");	
			}
   		 	if (arr != null && arr.length >= 1) {
   		 		return arr[0].trim();
   		 	}	
		} 
		return destination;
	}
	
	public String getDestinationCountry() {
		String destination = this.getDestination();
		if(destination != null) {
			String[] arr = null;
			if (destination.contains(",")) {
				arr = destination.split(",");	
			} else if  (destination.contains("-")) {
				arr = destination.split("-");	
			} else if  (destination.contains("–")) {
				arr = destination.split("–");	
			}
   		 	if (arr != null && arr.length >= 2) {
   		 		return arr[1].trim();
   		 	}	
		} 
		return destination;
	}

	
}






















