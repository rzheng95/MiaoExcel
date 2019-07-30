package com.rzheng.util;

public class Item {

	private String itemNumber;
	private String customerSkuNumber;
	private String description;
	private String htsCode;
	private String unitCube;
	private String quantity;
	private String unitCost;
	private String netAmount;
	
	public Item(String itemNumber, String customerSkuNumber, String description, String htsCode, String unitCube,
			String quantity, String unitCost, String netAmount) {
		super();
		this.itemNumber = itemNumber;
		this.customerSkuNumber = customerSkuNumber;
		this.description = description;
		this.htsCode = htsCode;
		this.unitCube = unitCube;
		this.quantity = quantity;
		this.unitCost = unitCost;
		this.netAmount = netAmount;
	}

	public String getItemNumber() {
		return itemNumber;
	}

	public void setItemNumber(String itemNumber) {
		this.itemNumber = itemNumber;
	}

	public String getCustomerSkuNumber() {
		return customerSkuNumber;
	}

	public void setCustomerSkuNumber(String customerSkuNumber) {
		this.customerSkuNumber = customerSkuNumber;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getHtsCode() {
		return htsCode;
	}

	public void setHtsCode(String htsCode) {
		this.htsCode = htsCode;
	}

	public String getUnitCube() {
		return unitCube;
	}

	public void setUnitCube(String unitCube) {
		this.unitCube = unitCube;
	}

	public String getQuantity() {
		return quantity;
	}

	public void setQuantity(String quantity) {
		this.quantity = quantity;
	}

	public String getUnitCost() {
		return unitCost;
	}

	public void setUnitCost(String unitCost) {
		this.unitCost = unitCost;
	}

	public String getNetAmount() {
		return netAmount;
	}

	public void setNetAmount(String netAmount) {
		this.netAmount = netAmount;
	}

	@Override
	public String toString() {
		return "Item [itemNumber=" + itemNumber + ", customerSkuNumber=" + customerSkuNumber + ", description="
				+ description + ", htsCode=" + htsCode + ", unitCube=" + unitCube + ", quantity=" + quantity
				+ ", unitCost=" + unitCost + ", netAmount=" + netAmount + "]";
	}
	
	
	
	
}
