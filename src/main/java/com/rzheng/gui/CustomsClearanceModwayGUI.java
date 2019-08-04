package com.rzheng.gui;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.UIManager;

import com.rzheng.magnussen.ShippingOrder;
import com.rzheng.modway.CustomsClearanceModway;

public class CustomsClearanceModwayGUI extends JFrame {

	private static final long serialVersionUID = 3152136175087201711L;
	
	// Title
	protected JLabel label_title;

	// PI
	protected JLabel label_proforma_invoice;
	protected JTextField textField_proforma_invoice;
	protected JButton button_proforma_invoice;
	
	// Ocean Bill of Lading
	protected JLabel label_ocean_bill_of_lading;
	protected JTextField textField_ocean_bill_of_lading;
	protected JButton button_ocean_bill_of_lading;
	
	// Product Dimension Chart
	protected JLabel label_product_dimension_chart;
	protected JTextField textField_product_dimension_chart;
	protected JButton button_product_dimension_chart;
	
	// Customs Clearance Template
	protected JLabel label_cc_template;
	protected JTextField textField_cc_template;
	protected JButton button_cc_template;
	

	// Output directory
	protected JLabel label_output_directory;
	protected JTextField textField_output_directory;
	protected JButton button_output_directory;
	
	// Invoice Number
	private JLabel label_invoice_number;
	private JTextField textField_invoice_number;
	
	// ETD
	private JLabel label_etd;
	private JTextField textField_etd;
	
	// ETA
	private JLabel label_eta;
	private JTextField textField_eta;
	
	// Generate Button
	protected JButton button_generate;
	
	// Required Textfields
	protected List<JTextField> requiredTextFields;
	
	protected int width = 880;
	protected int height = 560;
	public CustomsClearanceModwayGUI() {
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
//		String pi_path, String oceanBillOfLading_path, String product_dimension_chart_path, String cc_template, String cc_xls_path, String invoiceNumber, String etd, String eta
		
		this.requiredTextFields = new ArrayList<>();
		
		// Menu
		GUIFactory.createMenu(this);
		
		// Title
		label_title = setTitle(width);
		label_title.setFont(new Font("SimSun", Font.PLAIN, 30));
		add(label_title);
		
		// Pro Forma Invoice
		label_proforma_invoice = GUIFactory.createLabel("*PI: (Pro Forma Invoice)", 10, 80, 250, 20);
		add(label_proforma_invoice);

		textField_proforma_invoice = GUIFactory.createTextField(10, 100, 800, 23);
		add(textField_proforma_invoice);

		button_proforma_invoice = GUIFactory.createButton("...", 820, 100, 30, 23);
		add(button_proforma_invoice);

		button_proforma_invoice.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_proforma_invoice,
				"Select Pro Forma Invoice", "Excel Files", "OPEN", "xls"));
		
		// Ocean Bill of Lading
		label_ocean_bill_of_lading = GUIFactory.createLabel("*\u6d77\u8fd0\u63d0\u5355\u0028\u4ee3\u7406\u0029: (Ocean Bill of Lading)", 10, 130, 400, 20);
		add(label_ocean_bill_of_lading);

		textField_ocean_bill_of_lading = GUIFactory.createTextField(10, 150, 800, 23);
		add(textField_ocean_bill_of_lading);

		button_ocean_bill_of_lading = GUIFactory.createButton("...", 820, 150, 30, 23);
		add(button_ocean_bill_of_lading);

		button_ocean_bill_of_lading.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_ocean_bill_of_lading,
				"Select Ocean Bill of Lading", "PDF Files", "OPEN", "pdf"));
		
		// Product Dimension Chart
		label_product_dimension_chart = GUIFactory.createLabel("*\u5206\u8d27\u002d\u6709\u51c0\u6bdb\u4f53: (Product Dimension Chart)", 10, 180, 400, 20);
		add(label_product_dimension_chart);

		textField_product_dimension_chart = GUIFactory.createTextField(10, 200, 800, 23);
		add(textField_product_dimension_chart);

		button_product_dimension_chart = GUIFactory.createButton("...", 820, 200, 30, 23);
		add(button_product_dimension_chart);

		button_product_dimension_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_product_dimension_chart,
				"Select Product Dimension Chart", "Excel Files", "OPEN", "xls"));

		
		// Customs Clearance Template
		label_cc_template = GUIFactory.createLabel("*\u6e05\u5173\u6a21\u677f: (Modway Customs Clearance Template)", 10, 230, 350, 20);
		add(label_cc_template);

		textField_cc_template = GUIFactory.createTextField(10, 250, 800, 23);
		textField_cc_template.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\Modway Customs Clearance Template.xls");
		add(textField_cc_template);

		button_cc_template = GUIFactory.createButton("...", 820, 250, 30, 23);
		add(button_cc_template);

		button_cc_template.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_cc_template,
				"Select Customs Clearance Template", "Excel Files", "OPEN", "xls"));
		
		// Output Directory
		label_output_directory = GUIFactory.createLabel("\u5bfc\u51fa\u6587\u4ef6\u5939: (Output Directory)", 10, 280, 250, 20);
		add(label_output_directory);
	
		textField_output_directory = GUIFactory.createTextField(10, 300, 800, 23);
		add(textField_output_directory);
	
		button_output_directory = GUIFactory.createButton("...", 820, 300, 30, 23);
		add(button_output_directory);
	
		button_output_directory.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_output_directory,
				"Select Output Directory", "FOLDERS ONLY", "SAVE", "xls"));
		
		
		// Generate Button
		button_generate = GUIFactory.createButton("Generate Custom Clearance", 280, 400, 300, 50);
		add(button_generate);
		
		Calendar calendar = Calendar.getInstance();
	
		// 发票号 (Invoice Number)
		label_invoice_number = GUIFactory.createLabel("*\u53d1\u7968\u53f7: (Invoice Number)", 10, 330, 200, 23);
		add(label_invoice_number);
		
		textField_invoice_number = GUIFactory.createTextField(10, 350, 200, 23);
		textField_invoice_number.setText("INYB" + calendar.get(Calendar.YEAR) + "US");
		add(textField_invoice_number);
		
		// ETD (Estimated Time of Departure)
		label_etd = GUIFactory.createLabel("ETD: (Estimated Time of Departure)", 10, 380, 250, 23);
		add(label_etd);
		
		textField_etd = GUIFactory.createTextField(10, 400, 200, 23);
		add(textField_etd);
		
		// ETA (Estimated Time of Arrival)
		label_eta = GUIFactory.createLabel("ETA: (Estimated Time of Arrival)", 10, 430, 250, 23);
		add(label_eta);
		
		textField_eta = GUIFactory.createTextField(10, 450, 200, 23);
		add(textField_eta);
		
		requiredTextFields.add(textField_proforma_invoice);
		requiredTextFields.add(textField_ocean_bill_of_lading);
		requiredTextFields.add(textField_product_dimension_chart);
		requiredTextFields.add(textField_cc_template);
		requiredTextFields.add(textField_invoice_number);

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
	
	public void generate() {
		try {
			CustomsClearanceModway cc = new CustomsClearanceModway(
					textField_proforma_invoice.getText(), 
					textField_ocean_bill_of_lading.getText(), 
					textField_product_dimension_chart.getText(), 
					textField_cc_template.getText(), 
					textField_output_directory.getText(),
					textField_invoice_number.getText(),
					textField_etd.getText(),
					textField_eta.getText()
					);
			
			JOptionPane.showMessageDialog(null, cc.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
	
	public JLabel setTitle(int width) {
		return GUIFactory.createLabel("Custom Clearance", (width-255)/2, 5, 255, 80);
	}
}




















