package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Floor extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   */

  @DISPID(235)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   */

  @DISPID(128)
  @PropGet
  com.exceljava.com4j.excel.Border getBorder();


  /**
   */

  @DISPID(112)
  java.lang.Object clearFormats();


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   */

  @DISPID(129)
  @PropGet
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   */

  @DISPID(1663)
  @PropGet
  com.exceljava.com4j.excel.ChartFillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "PictureType"
   * </p>
   */

  @DISPID(161)
  @PropGet
  java.lang.Object getPictureType();


  /**
   * <p>
   * Setter method for the COM property "PictureType"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(161)
  @PropPut
  void setPictureType(
    java.lang.Object rhs);


  /**
   */

  @DISPID(211)
  void paste();


  /**
   * <p>
   * Getter method for the COM property "Thickness"
   * </p>
   */

  @DISPID(2652)
  @PropGet
  int getThickness();


  /**
   * <p>
   * Setter method for the COM property "Thickness"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2652)
  @PropPut
  void setThickness(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   */

  @DISPID(116)
  @PropGet
  com.exceljava.com4j.excel.ChartFormat getFormat();


  // Properties:
}
