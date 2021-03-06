package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002447C-0001-0000-C000-000000000046}")
public interface IXmlMaps extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter rootElementName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(schema, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param schema Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.XmlMap add(
    java.lang.String schema);

  /**
   * @param schema Mandatory java.lang.String parameter.
   * @param rootElementName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(10)
  com.exceljava.com4j.excel.XmlMap add(
    java.lang.String schema,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rootElementName);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.excel.XmlMap get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(12)
  com.exceljava.com4j.excel.XmlMap getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
