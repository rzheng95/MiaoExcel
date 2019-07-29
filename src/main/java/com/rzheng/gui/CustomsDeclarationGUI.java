package com.rzheng.gui;

import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import com.rzheng.excel.CustomsDeclaration;

public class CustomsDeclarationGUI extends ShippingOrderGUI {

	/**
	 * 
	 */
	private static final long serialVersionUID = -7643164752174547671L;

	private JLabel label_invoice_number;
	private JTextField textField_invoice_number;
	
	public CustomsDeclarationGUI() {
		super();
		label_shipping_order_template.setText("*\u6258\u4e66\u6a21\u677f: (Customs Declaration Template)");
		
		label_invoice_number = GUIFactory.createLabel("*PI Invoice Number:", 10, 380, 200, 23);
		add(label_invoice_number);
		
		textField_invoice_number = GUIFactory.createTextField(10, 400, 200, 23);
		add(textField_invoice_number);
		
	}
	
	@Override
	public void generate() {
		if (!textField_invoice_number.getText().isEmpty()) {
			try {
				CustomsDeclaration cd = new CustomsDeclaration(textField_product_chart.getText(),
						textField_dimension_chart.getText(), textField_shipping_instructions.getText(),
						textField_proforma_invoice.getText(), textField_output_directory.getText(),
						textField_shipping_order_template.getText(),
						textField_invoice_number.getText());
				cd.run();

			} catch (IOException e1) {
				e1.printStackTrace();
			}
		} else {
			JOptionPane.showMessageDialog(null, "* Textfields Are Requried.");
		}
	}
	
	@Override
	public JLabel setTitle(int screenWidth) {
		return GUIFactory.createLabel("Customs Declaration", (screenWidth-290)/2, 5, 290, 80);
	}
	
	@Override
	public JButton setGenerateButton() {
		return GUIFactory.createButton("Generate Customs Declaration", 280, 400, 300, 50);
	}
}













