package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1712-0000-0000-C000-000000000046}")
public interface Axes extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(type, 1);
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.XlAxisType parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoAxis
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.IMsoAxis getItem(
    com.exceljava.com4j.office.XlAxisType type);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoAxis
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.office.IMsoAxis getItem(
    com.exceljava.com4j.office.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.XlAxisGroup axisGroup);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(11)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(type, 1);
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.XlAxisType parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoAxis
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(12)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.IMsoAxis get_Default(
    com.exceljava.com4j.office.XlAxisType type);

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.XlAxisType parameter.
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoAxis
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.office.IMsoAxis get_Default(
    com.exceljava.com4j.office.XlAxisType type,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.XlAxisGroup axisGroup);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  // Properties:
}
