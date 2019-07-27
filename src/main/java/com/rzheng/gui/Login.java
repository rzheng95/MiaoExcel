package com.rzheng.gui;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.UIManager;

public class Login extends JFrame {
	
	private JButton button;
	public Login() {
		super("Login");
		try{UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());}catch(Exception e){}

		
		button = new JButton("Login");
		
		getContentPane().setLayout(null);
		setLocationRelativeTo(null);// put this after setSize and pack
		
		JLabel welcomMsg = new JLabel("Welcome! My \u732a\u732a  ");
		welcomMsg.setFont(new Font("SimSun",Font.PLAIN, 12));
		welcomMsg.setBounds(10, 130, 300, 20);
		getContentPane().add(welcomMsg);
		
		JLabel lblUsername = new JLabel("Username:");
		lblUsername.setFont(new Font("Tahoma", Font.PLAIN, 12));
		lblUsername.setBounds(10, 10, 80, 20);
		getContentPane().add(lblUsername);
		
		JTextField textFieldUsername = new JTextField();
		textFieldUsername.setFont(new Font("Tahoma", Font.PLAIN, 12));
		textFieldUsername.setBounds(10, 30, 290, 23);
		getContentPane().add(textFieldUsername);
		textFieldUsername.setColumns(10);
		
		JLabel lblPassword = new JLabel("Password:");
		lblPassword.setFont(new Font("Tahoma", Font.PLAIN, 12));
		lblPassword.setBounds(10, 60, 80, 20);
		getContentPane().add(lblPassword);
		
		
		
		final JPasswordField passwordField = new JPasswordField();
		passwordField.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ENTER ) {
					System.out.println(passwordField.getPassword());
					button.doClick();
				}
			}
		});
		passwordField.setBounds(10, 80, 290, 23);
		getContentPane().add(passwordField);
		
		button.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				dispose();
				new ShippingOrderGUI().setVisible(true);
			}
		});
		
		JMenuBar menuBar = new JMenuBar();
		setJMenuBar(menuBar);
		
		JMenu mnSelect = new JMenu("Select");
		menuBar.add(mnSelect);
		
		JMenuItem mntmExit = new JMenuItem("Exit");
		mntmExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});
		
		mnSelect.add(mntmExit);
		
		JMenu mnAbout = new JMenu("About");
		menuBar.add(mnAbout);
		
		add(button);
		
		setSize(330,250);
		Toolkit toolkit = Toolkit.getDefaultToolkit();  
		Dimension screenSize = toolkit.getScreenSize();
		// Calculate the frame location  
		int x = (screenSize.width - getWidth()) / 2;  
		int y = (screenSize.height - getHeight()) / 2;  
		setLocation(x, y); 
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setVisible(true);
	}
}
