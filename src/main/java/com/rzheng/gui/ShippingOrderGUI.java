package com.rzheng.gui;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.UIManager;

import com.rzheng.magnussen.ShippingOrder;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ShippingOrderGUI extends GUITemplate {

	private static final long serialVersionUID = 223134307400367242L;

	public ShippingOrderGUI() {
		textField_shipping_order_template.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\Shipping Order Template.xls");
	}
	
	@Override
	public void generate() {
		try {
			ShippingOrder so = new ShippingOrder(textField_product_chart.getText(), textField_dimension_chart.getText(), textField_shipping_instructions.getText(), textField_proforma_invoice.getText(), textField_output_directory.getText(), textField_shipping_order_template.getText());
			JOptionPane.showMessageDialog(null, so.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}

}





















