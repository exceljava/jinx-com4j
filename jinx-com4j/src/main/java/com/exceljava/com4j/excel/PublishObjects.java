package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PublishObjects extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter sheet is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter htmlType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter divID is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter htmlType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter divID is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, filename, sheet, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional java.lang.Object sheet);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter htmlType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter divID is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, filename, sheet, source, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional java.lang.Object sheet,
    @Optional java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter divID is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, filename, sheet, source, htmlType, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param htmlType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional java.lang.Object sheet,
    @Optional java.lang.Object source,
    @Optional java.lang.Object htmlType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, filename, sheet, source, htmlType, divID, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param htmlType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param divID Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional java.lang.Object sheet,
    @Optional java.lang.Object source,
    @Optional java.lang.Object htmlType,
    @Optional java.lang.Object divID);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param htmlType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param divID Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional java.lang.Object sheet,
    @Optional java.lang.Object source,
    @Optional java.lang.Object htmlType,
    @Optional java.lang.Object divID,
    @Optional java.lang.Object title);


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
  com.exceljava.com4j.excel.PublishObject getItem(
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
  com.exceljava.com4j.excel.PublishObject get_Default(
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
   */

  @DISPID(117)
  void delete();


  /**
   */

  @DISPID(1895)
  void publish();


  // Properties:
}
