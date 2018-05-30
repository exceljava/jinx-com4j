package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002447E-0001-0000-C000-000000000046}")
public interface IXPath extends Com4jObject {
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
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getValue();


  /**
   * <p>
   * Getter method for the COM property "Map"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(12)
  com.exceljava.com4j.excel.XmlMap getMap();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter selectionNamespace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter repeating is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue(map, xPath, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param xPath Mandatory java.lang.String parameter.
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void setValue(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String xPath);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter repeating is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setValue(map, xPath, selectionNamespace, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setValue(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespace);

  /**
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param repeating Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(13)
  void setValue(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespace,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object repeating);


  /**
   */

  @VTID(14)
  void clear();


  /**
   * <p>
   * Getter method for the COM property "Repeating"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getRepeating();


  // Properties:
}
