package com.rzheng.gui;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.UIManager;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.util.List;

public class ShippingOrderGUI extends JFrame {
	
	private List<Object> components;
	
	// Product Chart
	private JLabel label_product_chart;
	private JTextField textField_product_chart;
	private JButton button_product_chart;

	// Dimension Chart
	private JLabel label_dimension_chart;
	private JTextField textField_dimension_chart;
	private JButton button_dimension_chart;
	
	// Shipping Order Template
	private JLabel label_shipping_order_template;
	private JTextField textField_shipping_order_template;
	private JButton button_shipping_order_template;

	// SI
	private JLabel label_shipping_instructions;
	private JTextField textField_shipping_instructions;
	private JButton button_shipping_instructions;
	
	// PI
	private JLabel label_proforma_invoice;
	private JTextField textField_proforma_invoice;
	private JButton button_proforma_invoice;
	
	// Output directory
	private JLabel label_output_directory;
	private JTextField textField_output_directory;
	private JButton button_output_directory;
	
	// Generate Button
	private JButton button_generate;
	
	
	public ShippingOrderGUI() {
// DEFAULT SETTINGS --------------------------------------------------------------------------------------------
		super("Miao Excel");
		try{UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());}catch(Exception e){}
		getContentPane().setLayout(null);
		setLocationRelativeTo(null);// put this after setSize and pack
		int width = 880;
		int height = 600;
		setSize(width, height);
		Toolkit toolkit = Toolkit.getDefaultToolkit();  
		Dimension screenSize = toolkit.getScreenSize();
		// Calculate the frame location  
		int x = (screenSize.width - getWidth()) / 2;  
		int y = (screenSize.height - getHeight()) / 2;  
		setLocation(x, y); 
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setVisible(true);
// -------------------------------------------------------------------------------------------------------------
		
		// Menu
		GUIFactory.createMenu(this);
		
		// Product Chart
		label_product_chart = GUIFactory.createLabel("*\u4ea7\u54c1\u5bf9\u7167\u8868: (Product Chart)", 10, 10, 250, 20);
		add(label_product_chart);

		textField_product_chart = GUIFactory.createTextField(10, 30, 800, 23);
		add(textField_product_chart);

		button_product_chart = GUIFactory.createButton("...", 820, 30, 30, 23);
		add(button_product_chart);

		button_product_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_product_chart,
				"Select Product Chart", "Excel Files", "OPEN", "xls", "xlsx"));

		
		// Dimension Chart
		label_dimension_chart = GUIFactory.createLabel("*\u51c0\u6bdb\u4f53\u7edf\u8ba1\u8868: (Dimension Chart)", 10, 60, 250, 20);
		add(label_dimension_chart);

		textField_dimension_chart = GUIFactory.createTextField(10, 80, 800, 23);
		add(textField_dimension_chart);

		button_dimension_chart = GUIFactory.createButton("...", 820, 80, 30, 23);
		add(button_dimension_chart);

		button_dimension_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_dimension_chart,
				"Select Dimension Chart", "Excel Files", "OPEN", "xls", "xlsx"));
		
		// Shipping Order Template
		label_shipping_order_template = GUIFactory.createLabel("*\u6258\u4e66\u6a21\u677f: (Shipping Order Tempalte)", 10, 110, 250, 20);
		add(label_shipping_order_template);

		textField_shipping_order_template = GUIFactory.createTextField(10, 130, 800, 23);
		add(textField_shipping_order_template);

		button_shipping_order_template = GUIFactory.createButton("...", 820, 130, 30, 23);
		add(button_shipping_order_template);

		button_shipping_order_template.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_shipping_order_template,
				"Select Shipping Order Template", "Excel Files", "OPEN", "xls", "xlsx"));
		
		// SI
		label_shipping_instructions = GUIFactory.createLabel("*SI: (Shipping Instructions)", 10, 160, 250, 20);
		add(label_shipping_instructions);

		textField_shipping_instructions = GUIFactory.createTextField(10, 180, 800, 23);
		add(textField_shipping_instructions);

		button_shipping_instructions = GUIFactory.createButton("...", 820, 180, 30, 23);
		add(button_shipping_instructions);

		button_shipping_instructions.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_shipping_instructions,
				"Select Shipping Instruction", "PDF Files", "OPEN", "pdf"));
		
		
		// PI
		label_proforma_invoice = GUIFactory.createLabel("*PI: (Pro Forma Invoice)", 10, 210, 250, 20);
		add(label_proforma_invoice);

		textField_proforma_invoice = GUIFactory.createTextField(10, 230, 800, 23);
		add(textField_proforma_invoice);

		button_proforma_invoice = GUIFactory.createButton("...", 820, 230, 30, 23);
		add(button_proforma_invoice);

		button_proforma_invoice.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_proforma_invoice,
				"Select Proforma Invoice", "OPEN", "PDF Files", "pdf"));
		
		// Output Directory
		label_output_directory = GUIFactory.createLabel("\u5bfc\u51fa\u6587\u4ef6\u5939: (Output Directory)", 10, 260, 250, 20);
		add(label_output_directory);

		textField_output_directory = GUIFactory.createTextField(10, 280, 800, 23);
		add(textField_output_directory);

		button_output_directory = GUIFactory.createButton("...", 820, 280, 30, 23);
		add(button_output_directory);

		button_output_directory.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_output_directory,
				"Select Output Directory", "FOLDERS ONLY", "SAVE", "FOLDERS ONLY"));
		
		
		// Generate Button
		button_generate = GUIFactory.createButton("Generate Shipping Order", 280, 330, 300, 50);
		add(button_generate);
	
	}
	
	
	
	
}





















