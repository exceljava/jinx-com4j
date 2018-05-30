package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0334-0000-0000-C000-000000000046}")
public interface PropertyTests extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(index, 1033);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PropertyTest
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.PropertyTest getItem(
    int index);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PropertyTest
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.PropertyTest getItem(
    int index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter secondValue is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.office.MsoConnector parameter connector is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, condition, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param condition Mandatory com.exceljava.com4j.office.MsoCondition parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.office.MsoConnector.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1"})
  void add(
    java.lang.String name,
    com.exceljava.com4j.office.MsoCondition condition);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter secondValue is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.office.MsoConnector parameter connector is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, condition, value, com4j.Variant.getMissing(), 1);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param condition Mandatory com.exceljava.com4j.office.MsoCondition parameter.
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.office.MsoConnector.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1"})
  void add(
    java.lang.String name,
    com.exceljava.com4j.office.MsoCondition condition,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoConnector parameter connector is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, condition, value, secondValue, 1);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param condition Mandatory com.exceljava.com4j.office.MsoCondition parameter.
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @param secondValue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {com.exceljava.com4j.office.MsoConnector.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  void add(
    java.lang.String name,
    com.exceljava.com4j.office.MsoCondition condition,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object secondValue);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param condition Mandatory com.exceljava.com4j.office.MsoCondition parameter.
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @param secondValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param connector Optional parameter. Default value is 1
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  void add(
    java.lang.String name,
    com.exceljava.com4j.office.MsoCondition condition,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object secondValue,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoConnector connector);


  /**
   * @param index Mandatory int parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(12)
  void remove(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
