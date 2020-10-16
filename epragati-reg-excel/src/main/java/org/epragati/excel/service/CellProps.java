package org.epragati.excel.service;

import java.io.Serializable;
/**
 * 
 * @author krishnarjun.pampana
 *
 */
public class CellProps implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String valign;
	private String halign;
	private String colour;
	private String fontWeight;
	private String fieldValue;
	private Class<?> classType;
	private Integer index;
	

	/**
	 * @return the classType
	 */
	public Class<?> getClassType() {
		return classType;
	}
	public String getFieldValue() {
		return fieldValue;
	}
	public void setFieldValue(String fieldValue) {
		this.fieldValue = fieldValue;
	}
	/**
	 * @return the index
	 */
	public Integer getIndex() {
		return index;
	}
	
	/**
	 * @param classType the classType to set
	 */
	public void setClassType(Class<?> classType) {
		this.classType = classType;
	}
	/**
	 * @param index the index to set
	 */
	public void setIndex(Integer index) {
		this.index = index;
	}
	/**
	 * @return the valign
	 */
	public String getValign() {
		return valign;
	}
	/**
	 * @return the halign
	 */
	public String getHalign() {
		return halign;
	}
	/**
	 * @return the colour
	 */
	public String getColour() {
		return colour;
	}
	/**
	 * @return the fontWeight
	 */
	public String getFontWeight() {
		return fontWeight;
	}
	/**
	 * @param valign the valign to set
	 */
	public void setValign(String valign) {
		this.valign = valign;
	}
	/**
	 * @param halign the halign to set
	 */
	public void setHalign(String halign) {
		this.halign = halign;
	}
	/**
	 * @param colour the colour to set
	 */
	public void setColour(String colour) {
		this.colour = colour;
	}
	/**
	 * @param fontWeight the fontWeight to set
	 */
	public void setFontWeight(String fontWeight) {
		this.fontWeight = fontWeight;
	}
}
