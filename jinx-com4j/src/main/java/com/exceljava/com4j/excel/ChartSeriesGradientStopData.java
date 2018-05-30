package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ChartSeriesGradientStopData extends Com4jObject {
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
   * Getter method for the COM property "StopColor"
   * </p>
   */

  @DISPID(3267)
  @PropGet
  com.exceljava.com4j.excel.SeriesGradientStopColorFormat getStopColor();


  /**
   * <p>
   * Getter method for the COM property "StopPositionType"
   * </p>
   */

  @DISPID(3268)
  @PropGet
  com.exceljava.com4j.excel.XlGradientStopPositionType getStopPositionType();


  /**
   * <p>
   * Setter method for the COM property "StopPositionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlGradientStopPositionType parameter.
   */

  @DISPID(3268)
  @PropPut
  void setStopPositionType(
    com.exceljava.com4j.excel.XlGradientStopPositionType rhs);


  /**
   * <p>
   * Getter method for the COM property "StopValue"
   * </p>
   */

  @DISPID(3269)
  @PropGet
  java.lang.String getStopValue();


  /**
   * <p>
   * Setter method for the COM property "StopValue"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(3269)
  @PropPut
  void setStopValue(
    java.lang.String rhs);


  // Properties:
}
