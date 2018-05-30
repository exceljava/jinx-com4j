package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244FA-0001-0000-C000-000000000046}")
public interface IChartSeriesGradientStopData extends Com4jObject {
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
   * Getter method for the COM property "StopColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SeriesGradientStopColorFormat
   */

  @VTID(10)
  com.exceljava.com4j.excel.SeriesGradientStopColorFormat getStopColor();


  /**
   * <p>
   * Getter method for the COM property "StopPositionType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlGradientStopPositionType
   */

  @VTID(11)
  com.exceljava.com4j.excel.XlGradientStopPositionType getStopPositionType();


  /**
   * <p>
   * Setter method for the COM property "StopPositionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlGradientStopPositionType parameter.
   */

  @VTID(12)
  void setStopPositionType(
    com.exceljava.com4j.excel.XlGradientStopPositionType rhs);


  /**
   * <p>
   * Getter method for the COM property "StopValue"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getStopValue();


  /**
   * <p>
   * Setter method for the COM property "StopValue"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(14)
  void setStopValue(
    java.lang.String rhs);


  // Properties:
}
