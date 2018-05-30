package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Axes extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(type, 1);
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   */

  @DISPID(170)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Axis item(
    com.exceljava.com4j.excel.XlAxisType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   */

  @DISPID(170)
  com.exceljava.com4j.excel.Axis item(
    com.exceljava.com4j.excel.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(type, 1);
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   */

  @DISPID(0)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Axis _Default(
    com.exceljava.com4j.excel.XlAxisType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   */

  @DISPID(0)
  @DefaultMethod
  com.exceljava.com4j.excel.Axis _Default(
    com.exceljava.com4j.excel.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup);


  // Properties:
}
