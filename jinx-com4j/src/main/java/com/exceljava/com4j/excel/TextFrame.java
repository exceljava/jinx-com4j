package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface TextFrame extends Com4jObject {
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
   * Getter method for the COM property "MarginBottom"
   * </p>
   */

  @DISPID(1745)
  @PropGet
  float getMarginBottom();


  /**
   * <p>
   * Setter method for the COM property "MarginBottom"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(1745)
  @PropPut
  void setMarginBottom(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginLeft"
   * </p>
   */

  @DISPID(1746)
  @PropGet
  float getMarginLeft();


  /**
   * <p>
   * Setter method for the COM property "MarginLeft"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(1746)
  @PropPut
  void setMarginLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginRight"
   * </p>
   */

  @DISPID(1747)
  @PropGet
  float getMarginRight();


  /**
   * <p>
   * Setter method for the COM property "MarginRight"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(1747)
  @PropPut
  void setMarginRight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginTop"
   * </p>
   */

  @DISPID(1748)
  @PropGet
  float getMarginTop();


  /**
   * <p>
   * Setter method for the COM property "MarginTop"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(1748)
  @PropPut
  void setMarginTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  com.exceljava.com4j.office.MsoTextOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    com.exceljava.com4j.office.MsoTextOrientation rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * characters(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(603)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters characters();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * characters(start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters characters(
    @Optional java.lang.Object start);

  /**
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  com.exceljava.com4j.excel.Characters characters(
    @Optional java.lang.Object start,
    @Optional java.lang.Object length);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   */

  @DISPID(136)
  @PropGet
  com.exceljava.com4j.excel.XlHAlign getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlHAlign parameter.
   */

  @DISPID(136)
  @PropPut
  void setHorizontalAlignment(
    com.exceljava.com4j.excel.XlHAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   */

  @DISPID(137)
  @PropGet
  com.exceljava.com4j.excel.XlVAlign getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlVAlign parameter.
   */

  @DISPID(137)
  @PropPut
  void setVerticalAlignment(
    com.exceljava.com4j.excel.XlVAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoSize"
   * </p>
   */

  @DISPID(614)
  @PropGet
  boolean getAutoSize();


  /**
   * <p>
   * Setter method for the COM property "AutoSize"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(614)
  @PropPut
  void setAutoSize(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   */

  @DISPID(975)
  @PropGet
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(975)
  @PropPut
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoMargins"
   * </p>
   */

  @DISPID(1749)
  @PropGet
  boolean getAutoMargins();


  /**
   * <p>
   * Setter method for the COM property "AutoMargins"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1749)
  @PropPut
  void setAutoMargins(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalOverflow"
   * </p>
   */

  @DISPID(2922)
  @PropGet
  com.exceljava.com4j.excel.XlOartVerticalOverflow getVerticalOverflow();


  /**
   * <p>
   * Setter method for the COM property "VerticalOverflow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOartVerticalOverflow parameter.
   */

  @DISPID(2922)
  @PropPut
  void setVerticalOverflow(
    com.exceljava.com4j.excel.XlOartVerticalOverflow rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalOverflow"
   * </p>
   */

  @DISPID(2923)
  @PropGet
  com.exceljava.com4j.excel.XlOartHorizontalOverflow getHorizontalOverflow();


  /**
   * <p>
   * Setter method for the COM property "HorizontalOverflow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOartHorizontalOverflow parameter.
   */

  @DISPID(2923)
  @PropPut
  void setHorizontalOverflow(
    com.exceljava.com4j.excel.XlOartHorizontalOverflow rhs);


  // Properties:
}
