package com.rzheng.gui;

import java.io.IOException;
import java.util.Calendar;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import com.rzheng.magnussen.CustomsClearance;

public class CustomsClearanceGUI extends GUIMagnussenTemplate {
	
	private static final long serialVersionUID = 3014353196570362202L;
	
	private JLabel label_logistics_comfirmation_path;
	private JTextField textField_logistics_comfirmation_path;
	private JButton button_logistics_comfirmation_path;
	
	private JLabel label_logistics_companies;
	private JComboBox<String> comboBox_logistics_companies;
	
	private JLabel label_invoice_number;
	private JTextField textField_invoice_number;
	
	private JLabel label_container_number;
	private JTextField textField_container_number;
	
	private JLabel label_seal_number;
	private JTextField textField_seal_number;

	public CustomsClearanceGUI() {
		super();
		setSize(width, height+120);
		textField_shipping_order_template.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\Customs Clearance Template.xlsx");
		label_shipping_order_template.setText("*\u6e05\u5173\u6a21\u677f: (Customs Clearance Template)");
		
		button_generate.setText("Generate Customs Clearance");
		button_generate.setLocation(button_generate.getX(), button_generate.getY()+100);
		
		
		
		// logistics_company
		label_logistics_companies = GUIFactory.createLabel("\u6307\u4ee3:", 10, 380, 250, 20);
		add(label_logistics_companies);
		
		String[] companies = {"Unitex Logistics"};
		comboBox_logistics_companies = new JComboBox<String>(companies);
		comboBox_logistics_companies.setBounds(10, 400, 180, 23);
		add(comboBox_logistics_companies);
		
		// logistics_comfirmation_path
		label_logistics_comfirmation_path = GUIFactory.createLabel("\u9884\u914d: (Bay Plan)", 200, 380, 250, 20);
		add(label_logistics_comfirmation_path);

		textField_logistics_comfirmation_path = GUIFactory.createTextField(200, 400, 610, 23);
		add(textField_logistics_comfirmation_path);

		button_logistics_comfirmation_path = GUIFactory.createButton("...", 820, 400, 30, 23);
		add(button_logistics_comfirmation_path);

		if (comboBox_logistics_companies.getSelectedItem().toString().equalsIgnoreCase(companies[0])) {
			button_logistics_comfirmation_path.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_logistics_comfirmation_path,
					"Select Bay Plan", "Doc Files", "OPEN", "doc", "docx"));
		}
		
		// 发票号 (Invoice Number)
		label_invoice_number = GUIFactory.createLabel("*\u53d1\u7968\u53f7: (Invoice Number)", 10, 430, 250, 20);
		add(label_invoice_number);
		
		Calendar calendar = Calendar.getInstance();
		textField_invoice_number = GUIFactory.createTextField(10, 450, 200, 23);
		textField_invoice_number.setText("INYB" + calendar.get(Calendar.YEAR) + "US");
		add(textField_invoice_number);
		requiredTextFields.add(textField_invoice_number);
		
		// 箱封号 (Container Number)
		label_container_number = GUIFactory.createLabel("\u7bb1\u5c01\u53f7: (Container Number)", 10, 480, 250, 20);
		add(label_container_number);
		
		textField_container_number = GUIFactory.createTextField(10, 500, 200, 23);
		add(textField_container_number);
		
		// 铅封号 (Seal Number)
		label_seal_number = GUIFactory.createLabel("\u94c5\u5c01\u53f7: (Seal Number)", 10, 530, 250, 20);
		add(label_seal_number);
		
		textField_seal_number = GUIFactory.createTextField(10, 550, 200, 23);
		add(textField_seal_number);

		
	}
	
	@Override
	public JLabel setTitle(int width) {
		return GUIFactory.createLabel("Customs Clearance", (width-290)/2, 5, 290, 80);
	}
	
	@Override
	public void generate() {

		try {
			CustomsClearance cc = new CustomsClearance(
					textField_product_chart.getText(),
					textField_dimension_chart.getText(), 
					textField_shipping_instructions.getText(),
					textField_proforma_invoice.getText(), 
					comboBox_logistics_companies.getSelectedItem().toString(), // logistics_company 
					textField_logistics_comfirmation_path.getText(), // logistics_comfirmation_path 
					textField_output_directory.getText(), 
					textField_shipping_order_template.getText(), 
					textField_invoice_number.getText(), // invoiceNumber 
					textField_container_number.getText(), // containerNumber 
					textField_seal_number.getText()); // sealNumber
			JOptionPane.showMessageDialog(null, cc.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}

	}

}
