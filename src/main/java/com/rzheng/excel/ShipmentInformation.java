package com.rzheng.excel;

import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

public class ShipmentInformation {

	public static void main(String[] args) {
		ShipmentInformation si = new ShipmentInformation("052059 SI.pdf");
		System.out.println(si.getCpoNumber());
	}

	private String si_pdf_path;

	public ShipmentInformation(String si_pdf_path) {
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
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
	}

	public String getConsignee() {
		int i = 0;
		String consignee = null;
		String[] lines = this.readLines();

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
		int i = 0;
		String notifyParty = null;
		String[] lines = this.readLines();

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
		int i = 0;
		String shipToAddress = null;
		String[] lines = this.readLines();

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
		int i = 0;
		String forwarder = null;
		String[] lines = this.readLines();

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
		int i = 0;
		String[] lines = this.readLines();

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PORT_OF_LOADING)) {
				return lines[i].toUpperCase().substring(Constants.PORT_OF_LOADING.length()).trim();
			}
			i++;
		}
		return null;
	}

	public String getPortOfDischarge() {
		int i = 0;
		String[] lines = this.readLines();

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.PORT_OF_DISCHARGE)) {
				return lines[i].toUpperCase().substring(Constants.PORT_OF_DISCHARGE.length()).trim();
			}
			i++;
		}
		return null;
	}

	public String getDestination() {
		int i = 0;
		String[] lines = this.readLines();

		while (i < lines.length) {
			if (lines[i].toUpperCase().contains(Constants.DESTINATION)) {
				return lines[i].toUpperCase().substring(Constants.DESTINATION.length()).trim();
			}
			i++;
		}

		return null;
	}
	
	public String getPoNumber() {
		int i = 0;
		String[] lines = this.readLines();

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
		int i = 0;
		String[] lines = this.readLines();

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
		int i = 0;
		String[] lines = this.readLines();

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

}






















