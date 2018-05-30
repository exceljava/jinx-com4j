package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002444D-0000-0000-C000-000000000046}")
public interface CubeFields extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @DISPID(170) //= 0xaa. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.excel.CubeField getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.CubeField get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param caption Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @DISPID(2186) //= 0x88a. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.excel.CubeField addSet(
    java.lang.String name,
    java.lang.String caption);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter caption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getMeasure(attributeHierarchy, function, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param attributeHierarchy Mandatory java.lang.Object parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @DISPID(3089) //= 0xc11. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.CubeField getMeasure(
    @MarshalAs(NativeType.VARIANT) java.lang.Object attributeHierarchy,
    com.exceljava.com4j.excel.XlConsolidationFunction function);

  /**
   * @param attributeHierarchy Mandatory java.lang.Object parameter.
   * @param function Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   * @param caption Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @DISPID(3089) //= 0xc11. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.excel.CubeField getMeasure(
    @MarshalAs(NativeType.VARIANT) java.lang.Object attributeHierarchy,
    com.exceljava.com4j.excel.XlConsolidationFunction function,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object caption);


  // Properties:
}
