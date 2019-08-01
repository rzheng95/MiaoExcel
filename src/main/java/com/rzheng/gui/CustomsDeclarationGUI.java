package com.rzheng.gui;

import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import com.rzheng.magnussen.CustomsDeclaration;

public class CustomsDeclarationGUI extends GUITemplate {

	private static final long serialVersionUID = -7643164752174547671L;

	private JLabel label_invoice_number;
	private JTextField textField_invoice_number;
	
	public CustomsDeclarationGUI() {
		super();
		textField_shipping_order_template.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\Customs Declaration Template.xls");
		label_shipping_order_template.setText("*\u6258\u4e66\u6a21\u677f: (Customs Declaration Template)");
		
		label_invoice_number = GUIFactory.createLabel("*PI Invoice Number:", 10, 380, 200, 23);
		add(label_invoice_number);
		
		textField_invoice_number = GUIFactory.createTextField(10, 400, 200, 23);
		add(textField_invoice_number);
		
		button_generate.setText("Generate Customs Declaration");
		
		requiredTextFields.add(textField_invoice_number);
	}
	
	@Override
	public void generate() {

		try {
			CustomsDeclaration cd = new CustomsDeclaration(
					textField_product_chart.getText(),
					textField_dimension_chart.getText(), 
					textField_shipping_instructions.getText(),
					textField_proforma_invoice.getText(), 
					textField_output_directory.getText(),
					textField_shipping_order_template.getText(),
					textField_invoice_number.getText());
			
			JOptionPane.showMessageDialog(null, cd.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}

	}
	
	@Override
	public JLabel setTitle(int width) {
		return GUIFactory.createLabel("Customs Declaration", (width-290)/2, 5, 290, 80);
	}
	
}













