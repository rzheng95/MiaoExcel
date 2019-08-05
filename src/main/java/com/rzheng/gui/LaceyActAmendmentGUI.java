package com.rzheng.gui;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.HeadlessException;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.UIManager;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;

import com.rzheng.modway.CustomsClearanceModway;
import com.rzheng.modway.LaceyActAmendment;

public class LaceyActAmendmentGUI extends JFrame {

	private static final long serialVersionUID = -7769844774137306520L;
	
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
	
	// Lacey Act Template
	protected JLabel label_lacey_act_template;
	protected JTextField textField_lacey_act_template;
	protected JButton button_lacey_act_template;
	
	// Output directory
	protected JLabel label_output_directory;
	protected JTextField textField_output_directory;
	protected JButton button_output_directory;
	
	// ETA
	private JLabel label_eta;
	private JTextField textField_eta;
	
	// Generate Button
	protected JButton button_generate;
	
	// Required Textfields
	protected List<JTextField> requiredTextFields;

	protected int width = 880;
	protected int height = 480;
	public LaceyActAmendmentGUI() {
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
		label_title = GUIFactory.createLabel("Lacey Act Amendment", (width-290)/2, 5, 290, 80);
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
		label_lacey_act_template = GUIFactory.createLabel("*Lacey Act \u6a21\u677f: (Lacey Act Template)", 10, 180, 400, 20);
		add(label_lacey_act_template);

		textField_lacey_act_template = GUIFactory.createTextField(10, 200, 800, 23);
		add(textField_lacey_act_template);

		button_lacey_act_template = GUIFactory.createButton("...", 820, 200, 30, 23);
		add(button_lacey_act_template);

		button_lacey_act_template.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_lacey_act_template,
				"Select Lacey Act Template", "Docx Files", "OPEN", "docx"));
		
		
		// Output Directory
		label_output_directory = GUIFactory.createLabel("\u5bfc\u51fa\u6587\u4ef6\u5939: (Output Directory)", 10, 230, 350, 20);
		add(label_output_directory);
	
		textField_output_directory = GUIFactory.createTextField(10, 250, 800, 23);
		add(textField_output_directory);
	
		button_output_directory = GUIFactory.createButton("...", 820, 250, 30, 23);
		add(button_output_directory);
	
		button_output_directory.addActionListener(new GUIFactory.OpenFileActionListener(this, textField_output_directory,
				"Select Output Directory", "FOLDERS ONLY", "SAVE", "xls"));
		
		// ETA (Estimated Time of Arrival)
		label_eta = GUIFactory.createLabel("ETA: (Estimated Time of Arrival)", 10, 280, 250, 23);
		add(label_eta);
		
		textField_eta = GUIFactory.createTextField(10, 300, 200, 23);
		add(textField_eta);
		
		
		requiredTextFields.add(textField_proforma_invoice);
		requiredTextFields.add(textField_ocean_bill_of_lading);
		requiredTextFields.add(textField_lacey_act_template);

		// Generate Button
		button_generate = GUIFactory.createButton("Generate Lacey Act", 280, 330, 300, 50);
		add(button_generate);
		
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
			LaceyActAmendment laa = new LaceyActAmendment(
					textField_proforma_invoice.getText(), 
					textField_ocean_bill_of_lading.getText(), 
					textField_lacey_act_template.getText(), 
					textField_output_directory.getText(),
					textField_eta.getText()
					);
			
			JOptionPane.showMessageDialog(null, laa.run());
		} catch (IOException | HeadlessException | InvalidFormatException | XmlException e1) {
			e1.printStackTrace();
		}
	}
}



















