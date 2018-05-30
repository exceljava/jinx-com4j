package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface OLEFormat extends Com4jObject {
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
   */

  @DISPID(304)
  void activate();


  /**
   * <p>
   * Getter method for the COM property "Object"
   * </p>
   */

  @DISPID(1049)
  @PropGet
  com4j.Com4jObject getObject();


  /**
   * <p>
   * Getter method for the COM property "progID"
   * </p>
   */

  @DISPID(1523)
  @PropGet
  java.lang.String getProgID();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter verb is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * verb(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(606)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void verb();

  /**
   * @param verb Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(606)
  void verb(
    @Optional java.lang.Object verb);


  // Properties:
}
