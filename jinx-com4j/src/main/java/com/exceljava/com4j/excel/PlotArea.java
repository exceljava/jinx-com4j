package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PlotArea extends Com4jObject {
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
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    double rhs);


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
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(127)
  @PropPut
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(126)
  @PropPut
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "_InsideLeft"
   * </p>
   */

  @DISPID(2654)
  @PropGet
  double get_InsideLeft();


  /**
   * <p>
   * Getter method for the COM property "_InsideTop"
   * </p>
   */

  @DISPID(2655)
  @PropGet
  double get_InsideTop();


  /**
   * <p>
   * Getter method for the COM property "_InsideWidth"
   * </p>
   */

  @DISPID(2656)
  @PropGet
  double get_InsideWidth();


  /**
   * <p>
   * Getter method for the COM property "_InsideHeight"
   * </p>
   */

  @DISPID(2657)
  @PropGet
  double get_InsideHeight();


  /**
   * <p>
   * Getter method for the COM property "InsideLeft"
   * </p>
   */

  @DISPID(1667)
  @PropGet
  double getInsideLeft();


  /**
   * <p>
   * Setter method for the COM property "InsideLeft"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1667)
  @PropPut
  void setInsideLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InsideTop"
   * </p>
   */

  @DISPID(1668)
  @PropGet
  double getInsideTop();


  /**
   * <p>
   * Setter method for the COM property "InsideTop"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1668)
  @PropPut
  void setInsideTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InsideWidth"
   * </p>
   */

  @DISPID(1669)
  @PropGet
  double getInsideWidth();


  /**
   * <p>
   * Setter method for the COM property "InsideWidth"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1669)
  @PropPut
  void setInsideWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InsideHeight"
   * </p>
   */

  @DISPID(1670)
  @PropGet
  double getInsideHeight();


  /**
   * <p>
   * Setter method for the COM property "InsideHeight"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1670)
  @PropPut
  void setInsideHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  com.exceljava.com4j.excel.XlChartElementPosition getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartElementPosition parameter.
   */

  @DISPID(133)
  @PropPut
  void setPosition(
    com.exceljava.com4j.excel.XlChartElementPosition rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   */

  @DISPID(116)
  @PropGet
  com.exceljava.com4j.excel.ChartFormat getFormat();


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @DISPID(3253)
  void setProperty(
    java.lang.String id,
    java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(3256)
  java.lang.Object getProperty(
    java.lang.String id);


  // Properties:
}
