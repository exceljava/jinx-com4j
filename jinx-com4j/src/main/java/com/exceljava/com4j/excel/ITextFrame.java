package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002443D-0001-0000-C000-000000000046}")
public interface ITextFrame extends Com4jObject {
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
   * Getter method for the COM property "MarginBottom"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(10)
  float getMarginBottom();


  /**
   * <p>
   * Setter method for the COM property "MarginBottom"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(11)
  void setMarginBottom(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginLeft"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(12)
  float getMarginLeft();


  /**
   * <p>
   * Setter method for the COM property "MarginLeft"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(13)
  void setMarginLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginRight"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(14)
  float getMarginRight();


  /**
   * <p>
   * Setter method for the COM property "MarginRight"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(15)
  void setMarginRight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "MarginTop"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(16)
  float getMarginTop();


  /**
   * <p>
   * Setter method for the COM property "MarginTop"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(17)
  void setMarginTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTextOrientation
   */

  @VTID(18)
  com.exceljava.com4j.office.MsoTextOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   */

  @VTID(19)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Characters characters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);

  /**
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(20)
  com.exceljava.com4j.excel.Characters characters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object length);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlHAlign
   */

  @VTID(21)
  com.exceljava.com4j.excel.XlHAlign getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlHAlign parameter.
   */

  @VTID(22)
  void setHorizontalAlignment(
    com.exceljava.com4j.excel.XlHAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlVAlign
   */

  @VTID(23)
  com.exceljava.com4j.excel.XlVAlign getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlVAlign parameter.
   */

  @VTID(24)
  void setVerticalAlignment(
    com.exceljava.com4j.excel.XlVAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoSize"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(25)
  boolean getAutoSize();


  /**
   * <p>
   * Setter method for the COM property "AutoSize"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(26)
  void setAutoSize(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(27)
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(28)
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoMargins"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getAutoMargins();


  /**
   * <p>
   * Setter method for the COM property "AutoMargins"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(30)
  void setAutoMargins(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalOverflow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlOartVerticalOverflow
   */

  @VTID(31)
  com.exceljava.com4j.excel.XlOartVerticalOverflow getVerticalOverflow();


  /**
   * <p>
   * Setter method for the COM property "VerticalOverflow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOartVerticalOverflow parameter.
   */

  @VTID(32)
  void setVerticalOverflow(
    com.exceljava.com4j.excel.XlOartVerticalOverflow rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalOverflow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlOartHorizontalOverflow
   */

  @VTID(33)
  com.exceljava.com4j.excel.XlOartHorizontalOverflow getHorizontalOverflow();


  /**
   * <p>
   * Setter method for the COM property "HorizontalOverflow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOartHorizontalOverflow parameter.
   */

  @VTID(34)
  void setHorizontalOverflow(
    com.exceljava.com4j.excel.XlOartHorizontalOverflow rhs);


  // Properties:
}
