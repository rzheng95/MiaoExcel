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

import com.rzheng.modway.CustomsClearanceModway;
import com.rzheng.modway.CustomsDeclarationModway;

public class CustomsDeclarationModwayGUI extends JFrame {

	// Title
	protected JLabel label_title;

	// PI
	protected JLabel label_proforma_invoice;
	protected JTextField textField_proforma_invoice;
	protected JButton button_proforma_invoice;
	
	// Product Dimension Chart
	protected JLabel label_product_dimension_chart;
	protected JTextField textField_product_dimension_chart;
	protected JButton button_product_dimension_chart;
	
	// Customs Declaration Template
	protected JLabel label_cd_template;
	protected JTextField textField_cd_template;
	protected JButton button_cd_template;
	
	// Output directory
	protected JLabel label_output_directory;
	protected JTextField textField_output_directory;
	protected JButton button_output_directory;
	
	// Invoice Number
	private JLabel label_invoice_number;
	private JTextField textField_invoice_number;
	
	// Invoice Date
	private JLabel label_invoice_date;
	private JTextField textField_invoice_date;
	
	
	private static final long serialVersionUID = 8701914961303690880L;

	
	// Generate Button
	protected JButton button_generate;
	
	// Required Textfields
	protected List<JTextField> requiredTextFields;
	
	protected int width = 880;
	protected int height = 500;
	public CustomsDeclarationModwayGUI() {
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
		label_title = GUIFactory.createLabel("Custom Declaration", (width-270)/2, 5, 270, 80);
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
		
		
		// Product Dimension Chart
		label_product_dimension_chart = GUIFactory.createLabel("*\u5206\u8d27\u002d\u6709\u51c0\u6bdb\u4f53: (Product Dimension Chart)", 10, 130, 400, 20);
		add(label_product_dimension_chart);

		textField_product_dimension_chart = GUIFactory.createTextField(10, 150, 800, 23);
		add(textField_product_dimension_chart);

		button_product_dimension_chart = GUIFactory.createButton("...", 820, 150, 30, 23);
		add(button_product_dimension_chart);

		button_product_dimension_chart.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_product_dimension_chart,
				"Select Product Dimension Chart", "Excel Files", "OPEN", "xls"));
		
		// Customs Declaration Template
		label_cd_template = GUIFactory.createLabel("*\u6e05\u5173\u6a21\u677f: (Modway Customs Declaration Template)", 10, 180, 350, 20);
		add(label_cd_template);

		textField_cd_template = GUIFactory.createTextField(10, 200, 800, 23);
		textField_cd_template.setText("C:\\Users\\yibei\\Desktop\\程序\\表格模板\\Modway Customs Declaration Template.xls");
		add(textField_cd_template);

		button_cd_template = GUIFactory.createButton("...", 820, 200, 30, 23);
		add(button_cd_template);

		button_cd_template.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_cd_template,
				"Select Customs Declaration Template", "Excel Files", "OPEN", "xls"));
		
		// Output Directory
		label_output_directory = GUIFactory.createLabel("\u5bfc\u51fa\u6587\u4ef6\u5939: (Output Directory)", 10, 230, 250, 20);
		add(label_output_directory);
	
		textField_output_directory = GUIFactory.createTextField(10, 250, 800, 23);
		add(textField_output_directory);
	
		button_output_directory = GUIFactory.createButton("...", 820, 250, 30, 23);
		add(button_output_directory);
	
		button_output_directory.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_output_directory,
				"Select Output Directory", "FOLDERS ONLY", "SAVE", "xls"));
		
		Calendar calendar = Calendar.getInstance();
		
		// Invoice Number
		label_invoice_number = GUIFactory.createLabel("*\u53d1\u7968\u53f7: (Invoice Number)", 10, 280, 200, 23);
		add(label_invoice_number);
		
		textField_invoice_number = GUIFactory.createTextField(10, 300, 200, 23);
		textField_invoice_number.setText("INYB" + calendar.get(Calendar.YEAR) + "US");
		add(textField_invoice_number);
		
		// Invoice Date
		label_invoice_date = GUIFactory.createLabel("\u53d1\u7968\u65e5\u671f: (Invoice Date)", 10, 330, 200, 23);
		add(label_invoice_date);
		
		textField_invoice_date = GUIFactory.createTextField(10, 350, 200, 23);
		textField_invoice_date.setText(calendar.get(Calendar.YEAR) + "-");
		add(textField_invoice_date);
		
		
		requiredTextFields.add(textField_proforma_invoice);
		requiredTextFields.add(textField_product_dimension_chart);
		requiredTextFields.add(textField_cd_template);
		requiredTextFields.add(textField_invoice_number);
		
		// Generate Button
		button_generate = GUIFactory.createButton("Generate Custom Declaration", 280, 350, 300, 50);
		add(button_generate);

		button_generate.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {

				if (GUIFactory.isTextFieldEmpty(requiredTextFields)) {
					generate();
				} else {
					JOptionPane.showMessageDialog(null, "* Textfields Are Requried.");
				}
			}
		});
	}

	public void generate() {
		try {
			CustomsDeclarationModway cd = new CustomsDeclarationModway(
					textField_proforma_invoice.getText(),
					textField_product_dimension_chart.getText(), 
					textField_cd_template.getText(),
					textField_output_directory.getText(), 
					textField_invoice_number.getText(),
					textField_invoice_date.getText());

			JOptionPane.showMessageDialog(null, cd.run());
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
}












