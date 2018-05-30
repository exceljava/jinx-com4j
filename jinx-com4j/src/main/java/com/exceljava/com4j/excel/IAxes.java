package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002085B-0001-0000-C000-000000000046}")
public interface IAxes extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Axis
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Axis item(
    com.exceljava.com4j.excel.XlAxisType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.excel.Axis
   */

  @VTID(11)
  com.exceljava.com4j.excel.Axis item(
    com.exceljava.com4j.excel.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup);


  /**
   */

  @VTID(12)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Axis
   */

  @VTID(13)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Axis _Default(
    com.exceljava.com4j.excel.XlAxisType type);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.excel.Axis
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.excel.Axis _Default(
    com.exceljava.com4j.excel.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup);


  // Properties:
}
