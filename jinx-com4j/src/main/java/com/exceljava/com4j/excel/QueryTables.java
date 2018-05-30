package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface QueryTables extends Com4jObject,Iterable<Com4jObject> {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sql is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(connection, destination, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param connection Mandatory java.lang.Object parameter.
   * @param destination Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel._QueryTable add(
    java.lang.Object connection,
    com.exceljava.com4j.excel.Range destination);

  /**
   * @param connection Mandatory java.lang.Object parameter.
   * @param destination Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sql Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel._QueryTable add(
    java.lang.Object connection,
    com.exceljava.com4j.excel.Range destination,
    @Optional java.lang.Object sql);


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel._QueryTable item(
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
  com.exceljava.com4j.excel._QueryTable get_Default(
    java.lang.Object index);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
