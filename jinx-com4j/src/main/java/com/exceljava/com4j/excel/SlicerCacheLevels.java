package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SlicerCacheLevels extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter level is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(170)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SlicerCacheLevel getItem();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param level Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.SlicerCacheLevel getItem(
    @Optional java.lang.Object level);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter level is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SlicerCacheLevel get_Default();

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param level Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.SlicerCacheLevel get_Default(
    @Optional java.lang.Object level);


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
