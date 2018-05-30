package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{847EA60C-C1E6-4DC1-9847-78BC03A80AF0}")
public interface SeriesGradientStopData extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "StopColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SeriesGradientStopColorFormat
   */

  @VTID(8)
  com.exceljava.com4j.office.SeriesGradientStopColorFormat getStopColor();


  /**
   * <p>
   * Getter method for the COM property "StopPositionType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlGradientStopPositionType
   */

  @VTID(9)
  com.exceljava.com4j.office.XlGradientStopPositionType getStopPositionType();


  /**
   * <p>
   * Setter method for the COM property "StopPositionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.XlGradientStopPositionType parameter.
   */

  @VTID(10)
  void setStopPositionType(
    com.exceljava.com4j.office.XlGradientStopPositionType rhs);


  /**
   * <p>
   * Getter method for the COM property "StopValue"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getStopValue();


  /**
   * <p>
   * Setter method for the COM property "StopValue"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(12)
  void setStopValue(
    java.lang.String rhs);


  // Properties:
}
