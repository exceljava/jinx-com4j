package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Outline extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AutomaticStyles"
   * </p>
   */

  @DISPID(959)
  @PropGet
  boolean getAutomaticStyles();


  /**
   * <p>
   * Setter method for the COM property "AutomaticStyles"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(959)
  @PropPut
  void setAutomaticStyles(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowLevels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnLevels is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showLevels(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(960)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object showLevels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnLevels is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showLevels(rowLevels, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowLevels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(960)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object showLevels(
    @Optional java.lang.Object rowLevels);

  /**
   * @param rowLevels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnLevels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(960)
  java.lang.Object showLevels(
    @Optional java.lang.Object rowLevels,
    @Optional java.lang.Object columnLevels);


  /**
   * <p>
   * Getter method for the COM property "SummaryColumn"
   * </p>
   */

  @DISPID(961)
  @PropGet
  com.exceljava.com4j.excel.XlSummaryColumn getSummaryColumn();


  /**
   * <p>
   * Setter method for the COM property "SummaryColumn"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSummaryColumn parameter.
   */

  @DISPID(961)
  @PropPut
  void setSummaryColumn(
    com.exceljava.com4j.excel.XlSummaryColumn rhs);


  /**
   * <p>
   * Getter method for the COM property "SummaryRow"
   * </p>
   */

  @DISPID(902)
  @PropGet
  com.exceljava.com4j.excel.XlSummaryRow getSummaryRow();


  /**
   * <p>
   * Setter method for the COM property "SummaryRow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSummaryRow parameter.
   */

  @DISPID(902)
  @PropPut
  void setSummaryRow(
    com.exceljava.com4j.excel.XlSummaryRow rhs);


  // Properties:
}
