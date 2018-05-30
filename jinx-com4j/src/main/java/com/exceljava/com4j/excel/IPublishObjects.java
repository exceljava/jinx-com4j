package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024443-0001-0000-C000-000000000046}")
public interface IPublishObjects extends Com4jObject,Iterable<Com4jObject> {
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheet);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheet,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheet,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object htmlType);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheet,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object htmlType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object divID);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlSourceType parameter.
   * @param filename Mandatory java.lang.String parameter.
   * @param sheet Optional parameter. Default value is com4j.Variant.getMissing()
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param htmlType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param divID Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(10)
  com.exceljava.com4j.excel.PublishObject add(
    com.exceljava.com4j.excel.XlSourceType sourceType,
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheet,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object htmlType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object divID,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(12)
  com.exceljava.com4j.excel.PublishObject getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObject
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.excel.PublishObject get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   */

  @VTID(15)
  void delete();


  /**
   */

  @VTID(16)
  void publish();


  // Properties:
}
