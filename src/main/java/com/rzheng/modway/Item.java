package com.rzheng.modway;

public class Item {
	private String partNum;
	private String description;
	private String itemNum;
	private String fabric_leather;
	private int quantity;
	private double unitPrice;
	private double totalAmount;
	private double netWeight;
	private double grossWeight;
	private double cbm;
	

	


	public Item(String partNum, String description, String itemNum, String fabric_leather, int quantity,
			double unitPrice, double totalAmount) {
		super();
		this.partNum = partNum;
		this.description = description;
		this.itemNum = itemNum;
		this.fabric_leather = fabric_leather;
		this.quantity = quantity;
		this.unitPrice = unitPrice;
		this.totalAmount = totalAmount;
	}



	public Item(String partNum, String description, String itemNum, int quantity, double netWeight, double grossWeight,
			double cbm) {
		super();
		this.partNum = partNum;
		this.description = description;
		this.itemNum = itemNum;
		this.quantity = quantity;
		this.netWeight = netWeight;
		this.grossWeight = grossWeight;
		this.cbm = cbm;
	}



	public String getPartNum() {
		return partNum;
	}
	
	public String getStyleNum() {
		return partNum;
	}

	public void setPartNum(String partNum) {
		this.partNum = partNum;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getItemNum() {
		return itemNum;
	}
	
	public String getVendorStyleNum() {
		return itemNum;
	}

	public void setItemNum(String itemNum) {
		this.itemNum = itemNum;
	}

	public String getFabric_leather() {
		return fabric_leather;
	}

	public void setFabric_leather(String fabric_leather) {
		this.fabric_leather = fabric_leather;
	}

	public int getQuantity() {
		return quantity;
	}

	public void setQuantity(int quantity) {
		this.quantity = quantity;
	}

	public double getUnitPrice() {
		return unitPrice;
	}

	public void setUnitPrice(double unitPrice) {
		this.unitPrice = unitPrice;
	}

	public double getTotalAmount() {
		return totalAmount;
	}

	public void setTotalAmount(double totalAmount) {
		this.totalAmount = totalAmount;
	}

	public double getNetWeight() {
		return netWeight;
	}

	public void setNetWeight(double netWeight) {
		this.netWeight = netWeight;
	}

	public double getGrossWeight() {
		return grossWeight;
	}

	public void setGrossWeight(double grossWeight) {
		this.grossWeight = grossWeight;
	}

	public double getCbm() {
		return cbm;
	}

	public void setCbm(double cbm) {
		this.cbm = cbm;
	}

	@Override
	public String toString() {
		return "Item [partNum=" + partNum + ", description=" + description + ", itemNum=" + itemNum
				+ ", fabric_leather=" + fabric_leather + ", quantity=" + quantity + ", unitPrice=" + unitPrice
				+ ", totalAmount=" + totalAmount + ", netWeight=" + netWeight + ", grossWeight=" + grossWeight
				+ ", cmb=" + cbm + "]";
	}
	
	
	
	
}
