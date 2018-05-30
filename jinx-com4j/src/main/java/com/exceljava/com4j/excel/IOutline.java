package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208AB-0001-0000-C000-000000000046}")
public interface IOutline extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AutomaticStyles"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getAutomaticStyles();


  /**
   * <p>
   * Setter method for the COM property "AutomaticStyles"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object showLevels(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowLevels);

  /**
   * @param rowLevels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnLevels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object showLevels(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowLevels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnLevels);


  /**
   * <p>
   * Getter method for the COM property "SummaryColumn"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSummaryColumn
   */

  @VTID(13)
  com.exceljava.com4j.excel.XlSummaryColumn getSummaryColumn();


  /**
   * <p>
   * Setter method for the COM property "SummaryColumn"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSummaryColumn parameter.
   */

  @VTID(14)
  void setSummaryColumn(
    com.exceljava.com4j.excel.XlSummaryColumn rhs);


  /**
   * <p>
   * Getter method for the COM property "SummaryRow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSummaryRow
   */

  @VTID(15)
  com.exceljava.com4j.excel.XlSummaryRow getSummaryRow();


  /**
   * <p>
   * Setter method for the COM property "SummaryRow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSummaryRow parameter.
   */

  @VTID(16)
  void setSummaryRow(
    com.exceljava.com4j.excel.XlSummaryRow rhs);


  // Properties:
}
