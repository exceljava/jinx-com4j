package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotTables extends Com4jObject,Iterable<Com4jObject> {
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
  com.exceljava.com4j.excel.PivotTable item(
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
     * <li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter defaultVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(pivotCache, tableDestination, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pivotCache Mandatory com.exceljava.com4j.excel.PivotCache parameter.
   * @param tableDestination Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotTable add(
    com.exceljava.com4j.excel.PivotCache pivotCache,
    java.lang.Object tableDestination);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter defaultVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(pivotCache, tableDestination, tableName, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pivotCache Mandatory com.exceljava.com4j.excel.PivotCache parameter.
   * @param tableDestination Mandatory java.lang.Object parameter.
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotTable add(
    com.exceljava.com4j.excel.PivotCache pivotCache,
    java.lang.Object tableDestination,
    @Optional java.lang.Object tableName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter defaultVersion is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(pivotCache, tableDestination, tableName, readData, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pivotCache Mandatory com.exceljava.com4j.excel.PivotCache parameter.
   * @param tableDestination Mandatory java.lang.Object parameter.
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotTable add(
    com.exceljava.com4j.excel.PivotCache pivotCache,
    java.lang.Object tableDestination,
    @Optional java.lang.Object tableName,
    @Optional java.lang.Object readData);

  /**
   * @param pivotCache Mandatory com.exceljava.com4j.excel.PivotCache parameter.
   * @param tableDestination Mandatory java.lang.Object parameter.
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param defaultVersion Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.PivotTable add(
    com.exceljava.com4j.excel.PivotCache pivotCache,
    java.lang.Object tableDestination,
    @Optional java.lang.Object tableName,
    @Optional java.lang.Object readData,
    @Optional java.lang.Object defaultVersion);


  // Properties:
}
