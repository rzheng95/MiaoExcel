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

public class GUIMagnussenTemplate extends JFrame {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 425106102966770263L;
	
	
	// Title
	protected JLabel label_title;
	
	// Product Chart
	protected JLabel label_product_chart;
	protected JTextField textField_product_chart;
	protected JButton button_product_chart;

	// Dimension Chart
	protected JLabel label_dimension_chart;
	protected JTextField textField_dimension_chart;
	protected JButton button_dimension_chart;
	
	// Shipping Order Template
	protected JLabel label_shipping_order_template;
	protected JTextField textField_shipping_order_template;
	protected JButton button_shipping_order_template;

	// SI
	protected JLabel label_shipping_instructions;
	protected JTextField textField_shipping_instructions;
	protected JButton button_shipping_instructions;
	
	// PI
	protected JLabel label_proforma_invoice;
	protected JTextField textField_proforma_invoice;
	protected JButton button_proforma_invoice;
	
	// Output directory
	protected JLabel label_output_directory;
	protected JTextField textField_output_directory;
	protected JButton button_output_directory;
	
	// Generate Button
	protected JButton button_generate;
	
	// Required Textfields
	protected List<JTextField> requiredTextFields;
	
	protected int width = 880;
	protected int height = 530;
	
	public GUIMagnussenTemplate() {
// DEFAULT SETTINGS --------------------------------------------------------------------------------------------
		super("Miao Excel");
		try{UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());}catch(Exception e){}
		getContentPane().setLayout(null);
		setLocationRelativeTo(null);// put this after setSize and pack
		
		setSize(width, height);
		Toolkit toolkit = Toolkit.getDefaultToolkit();  
		Dimension screenSize = toolkit.getScreenSize();
		// Calculate the frame location  
		int x = (screenSize.width - getWidth()) / 2;  
		int y = (screenSize.height - getHeight()) / 2;  
		setLocation(x, y); 
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setResizable(false);
		setVisible(true);
// -------------------------------------------------------------------------------------------------------------
		
		this.requiredTextFields = new ArrayList<>();
		
		// Menu
		GUIFactory.createMenu(this);
		
		// Title
		label_title = setTitle(width);
		label_title.setFont(new Font("SimSun", Font.PLAIN, 30));
		add(label_title);
		
		
		// Product Chart
		label_product_chart = GUIFactory.createLabel("*\u4ea7\u54c1\u5bf9\u7167\u8868: (Product Chart)", 10, 80, 250, 20);
		add(label_product_chart);

		textField_product_chart = GUIFactory.createTextField(10, 100, 800, 23);
		textField_product_chart.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\magnussen 产品对照表 201905025.xlsx");
		add(textField_product_chart);

		button_product_chart = GUIFactory.createButton("...", 820, 100, 30, 23);
		add(button_product_chart);

		button_product_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_product_chart,
				"Select Product Chart", "Excel Files", "OPEN", "xls", "xlsx"));

		
		// Dimension Chart
		label_dimension_chart = GUIFactory.createLabel("*\u51c0\u6bdb\u4f53\u7edf\u8ba1\u8868: (Dimension Chart)", 10, 130, 250, 20);
		add(label_dimension_chart);

		textField_dimension_chart = GUIFactory.createTextField(10, 150, 800, 23);
		textField_dimension_chart.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\净毛体统计2016.09.07.xls");
		add(textField_dimension_chart);

		button_dimension_chart = GUIFactory.createButton("...", 820, 150, 30, 23);
		add(button_dimension_chart);

		button_dimension_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_dimension_chart,
				"Select Dimension Chart", "Excel Files", "OPEN", "xls", "xlsx"));
		
		// Shipping Order Template
		label_shipping_order_template = GUIFactory.createLabel("*\u6258\u4e66\u6a21\u677f: (Shipping Order Template)", 10, 180, 350, 20);
		add(label_shipping_order_template);

		textField_shipping_order_template = GUIFactory.createTextField(10, 200, 800, 23);
		add(textField_shipping_order_template);

		button_shipping_order_template = GUIFactory.createButton("...", 820, 200, 30, 23);
		add(button_shipping_order_template);

		button_shipping_order_template.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_shipping_order_template,
				"Select Shipping Order Template", "Excel Files", "OPEN", "xls", "xlsx"));
		
		// SI
		label_shipping_instructions = GUIFactory.createLabel("*SI: (Shipping Instructions)", 10, 230, 250, 20);
		add(label_shipping_instructions);

		textField_shipping_instructions = GUIFactory.createTextField(10, 250, 800, 23);
		add(textField_shipping_instructions);

		button_shipping_instructions = GUIFactory.createButton("...", 820, 250, 30, 23);
		add(button_shipping_instructions);

		button_shipping_instructions.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_shipping_instructions,
				"Select Shipping Instruction", "PDF Files", "OPEN", "pdf"));
		
		
		// PI
		label_proforma_invoice = GUIFactory.createLabel("*PI: (Pro Forma Invoice)", 10, 280, 250, 20);
		add(label_proforma_invoice);

		textField_proforma_invoice = GUIFactory.createTextField(10, 300, 800, 23);
		add(textField_proforma_invoice);

		button_proforma_invoice = GUIFactory.createButton("...", 820, 300, 30, 23);
		add(button_proforma_invoice);

		button_proforma_invoice.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_proforma_invoice,
				"Select Proforma Invoice", "OPEN", "PDF Files", "pdf"));
		
		// Output Directory
		label_output_directory = GUIFactory.createLabel("\u5bfc\u51fa\u6587\u4ef6\u5939: (Output Directory)", 10, 330, 250, 20);
		add(label_output_directory);

		textField_output_directory = GUIFactory.createTextField(10, 350, 800, 23);
		add(textField_output_directory);

		button_output_directory = GUIFactory.createButton("...", 820, 350, 30, 23);
		add(button_output_directory);

		button_output_directory.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_output_directory,
				"Select Output Directory", "FOLDERS ONLY", "SAVE", "xls"));
		
		
		// Generate Button
		button_generate = GUIFactory.createButton("Generate Shipping Order", 280, 400, 300, 50);
		add(button_generate);
	
		
		requiredTextFields.add(textField_product_chart);
		requiredTextFields.add(textField_dimension_chart);
		requiredTextFields.add(textField_shipping_order_template);
		requiredTextFields.add(textField_shipping_instructions);
		requiredTextFields.add(textField_proforma_invoice);

		button_generate.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				if(GUIFactory.isTextFieldEmpty(requiredTextFields)) {
					generate();
				} else {
					JOptionPane.showMessageDialog(null, "* Textfields Are Requried.");
				}
			}
		});
	}
	
	public JLabel setTitle(int width) {
		return GUIFactory.createLabel("Shipping Order", (width-210)/2, 5, 210, 80);
	}
	
	public void generate() {
		try {
			ShippingOrder so = new ShippingOrder(
					textField_product_chart.getText(), 
					textField_dimension_chart.getText(), 
					textField_shipping_instructions.getText(), 
					textField_proforma_invoice.getText(), 
					textField_output_directory.getText(), 
					textField_shipping_order_template.getText());
			
			JOptionPane.showMessageDialog(null, so.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
	

}





















