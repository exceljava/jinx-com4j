package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotCaches extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel.PivotCache item(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.PivotCache get_Default(
    java.lang.Object index);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotCache add(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.PivotCache add(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional java.lang.Object sourceData);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter version is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * create(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   */

  @DISPID(1896)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter version is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * create(sourceType, sourceData, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1896)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional java.lang.Object sourceData);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param version Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1896)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional java.lang.Object sourceData,
    @Optional java.lang.Object version);


  // Properties:
}
