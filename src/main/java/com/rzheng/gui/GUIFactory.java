package com.rzheng.gui;

import java.awt.Component;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;
import java.util.prefs.Preferences;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

public final class GUIFactory {

	public static JLabel createLabel(String text, int x, int y, int width, int height) {
		JLabel label = new JLabel(text);
		label.setFont(new Font("SimSun", Font.PLAIN, 12));
		label.setBounds(x, y, width, height);
		return label;
	}

	public static JTextField createTextField(int x, int y, int width, int height) {
		JTextField textfield = new JTextField();
		textfield.setFont(new Font("SimSun", Font.PLAIN, 12));
		textfield.setBounds(x, y, width, height);
		textfield.setColumns(10);
		return textfield;
	}

	public static JButton createButton(String text, int x, int y, int width, int height) {
		JButton button = new JButton(text);
		button.setFont(new Font("SimSun", Font.PLAIN, 12));
		button.setBounds(x, y, width, height);
		return button;
	}

	public static void createMenu(final JFrame frame) {
		JMenuBar menuBar = new JMenuBar();
		frame.setJMenuBar(menuBar);
		
		JMenu mnModway = new JMenu("Modway");
		menuBar.add(mnModway);

		JMenu mnMagnussen = new JMenu("Magnussen");
		menuBar.add(mnMagnussen);

		JMenuItem menuItem_shippingOrder = new JMenuItem("Shipping Order");
		menuItem_shippingOrder.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				new ShippingOrderGUI().setVisible(true);
			}
		});
		
		JMenuItem menuItem_customsDeclaration = new JMenuItem("Customs Declaration");
		menuItem_customsDeclaration.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				new CustomsDeclarationGUI().setVisible(true);
			}
		});
		
		JMenuItem menuItem_customsClearance = new JMenuItem("Customs Cleanrance");
		menuItem_customsClearance.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				new CustomsClearanceGUI().setVisible(true);
			}
		});
		

//		JMenuItem menuItem_exit = new JMenuItem("Exit");
//		menuItem_exit.addActionListener(new ActionListener() {
//			public void actionPerformed(ActionEvent e) {
//				System.exit(0);
//			}
//		});
//		mnMagnussen.add(menuItem_exit);
		
		mnMagnussen.add(menuItem_shippingOrder);
		mnMagnussen.add(menuItem_customsDeclaration);
		mnMagnussen.add(menuItem_customsClearance);
		
		
		JMenuItem menuItem_customsClearanceModway = new JMenuItem("Customs Cleanrance");
		menuItem_customsClearanceModway.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				new CustomsClearanceModwayGUI().setVisible(true);
			}
		});
		
		mnModway.add(menuItem_customsClearanceModway);
		

		JMenu mnAbout = new JMenu("About");
		JMenuItem menuItem_readMe = new JMenuItem("README");
		menuItem_readMe.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null,
						"This program is solely made by Richard Zheng for \u82d7\u9752's personal use in July 2019.\n\n"
						+ "No part of this publication may be reproduced, distributed, \n"
						+ "or transmitted in any form or by any means, including photocopying, \n"
						+ "recording, or other electronic or mechanical methods, \n"
						+ "without the prior written permission of the publisher, \n"
						+ "except in the case of brief quotations embodied in critical reviews and \n"
						+ "certain other noncommercial uses permitted by copyright law. \n"
						+ "All Rights Reserved.");
			}
		});

		mnAbout.add(menuItem_readMe);

		menuBar.add(mnAbout);

	}
	
	public static boolean isTextFieldEmpty(List<JTextField> textfields) {
		
		for (JTextField tf : textfields) {
			if (tf.getText().isEmpty())
				return false;
		}
		return true;
	}
	private static Preferences prefs;
	private static final String LAST_USED_FOLDER = "LAST_USED_FOLDER";
	public static class OpenFileActionListener implements ActionListener {

		private String title;
		private String description;
		private String[] extentions;
		private String mode;
		private JFileChooser chooser;
		private Component parent;
		private JTextField textfield;
	


		public OpenFileActionListener(Component parent, JTextField textfield, String title, String description,
				String mode, String... extentions) {
			this.title = title;
			this.description = description;
			this.mode = mode;
			this.extentions = extentions;
			this.parent = parent;
			this.textfield = textfield;
			prefs = Preferences.userRoot().node(getClass().getName());
			this.chooser = new JFileChooser();
		}

		@Override
		public void actionPerformed(ActionEvent e) {
			chooser.setCurrentDirectory(new File(prefs.get(LAST_USED_FOLDER, new File(".").getAbsolutePath())));
		
			int response = -1;
			chooser.setDialogTitle(title);
			if (mode.equalsIgnoreCase("SAVE")) {
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);
				response = chooser.showSaveDialog(parent);
			} else {
				FileFilter filter = new FileNameExtensionFilter(description, extentions);
				chooser.setFileFilter(filter);
				response = chooser.showOpenDialog(parent);
			}

			if (response == JFileChooser.APPROVE_OPTION) {
				prefs.put(LAST_USED_FOLDER, chooser.getSelectedFile().getParent());
				File file = chooser.getSelectedFile();
				textfield.setText(file.getPath());
			}
		}

	}
}
