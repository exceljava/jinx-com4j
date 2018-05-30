package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ConditionValue extends Com4jObject {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newvalue is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(newtype, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param newtype Mandatory com.exceljava.com4j.excel.XlConditionValueTypes parameter.
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void modify(
    com.exceljava.com4j.excel.XlConditionValueTypes newtype);

  /**
   * @param newtype Mandatory com.exceljava.com4j.excel.XlConditionValueTypes parameter.
   * @param newvalue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  void modify(
    com.exceljava.com4j.excel.XlConditionValueTypes newtype,
    @Optional java.lang.Object newvalue);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.excel.XlConditionValueTypes getType();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.Object getValue();


  // Properties:
}
