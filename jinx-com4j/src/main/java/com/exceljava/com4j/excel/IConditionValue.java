package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024492-0001-0000-C000-000000000046}")
public interface IConditionValue extends Com4jObject {
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

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void modify(
    com.exceljava.com4j.excel.XlConditionValueTypes newtype);

  /**
   * @param newtype Mandatory com.exceljava.com4j.excel.XlConditionValueTypes parameter.
   * @param newvalue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  void modify(
    com.exceljava.com4j.excel.XlConditionValueTypes newtype,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newvalue);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlConditionValueTypes
   */

  @VTID(11)
  com.exceljava.com4j.excel.XlConditionValueTypes getType();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue();


  // Properties:
}
