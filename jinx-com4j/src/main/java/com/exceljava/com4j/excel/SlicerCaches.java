package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SlicerCaches extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.SlicerCache getItem(
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
  com.exceljava.com4j.excel.SlicerCache get_Default(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, sourceField, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SlicerCache add(
    java.lang.Object source,
    java.lang.Object sourceField);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.SlicerCache add(
    java.lang.Object source,
    java.lang.Object sourceField,
    @Optional java.lang.Object name);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter slicerCacheType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(source, sourceField, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SlicerCache add2(
    java.lang.Object source,
    java.lang.Object sourceField);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter slicerCacheType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(source, sourceField, name, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.SlicerCache add2(
    java.lang.Object source,
    java.lang.Object sourceField,
    @Optional java.lang.Object name);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param slicerCacheType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  com.exceljava.com4j.excel.SlicerCache add2(
    java.lang.Object source,
    java.lang.Object sourceField,
    @Optional java.lang.Object name,
    @Optional java.lang.Object slicerCacheType);


  // Properties:
}
