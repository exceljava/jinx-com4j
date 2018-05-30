package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SortFields extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter sortOn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional java.lang.Object sortOn);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, order, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional java.lang.Object sortOn,
    @Optional java.lang.Object order);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, order, customOrder, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customOrder Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional java.lang.Object sortOn,
    @Optional java.lang.Object order,
    @Optional java.lang.Object customOrder);

  /**
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataOption Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional java.lang.Object sortOn,
    @Optional java.lang.Object order,
    @Optional java.lang.Object customOrder,
    @Optional java.lang.Object dataOption);


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.SortField getItem(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   */

  @DISPID(111)
  void clear();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.SortField get_Default(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
