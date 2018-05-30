package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244B0-0001-0000-C000-000000000046}")
public interface IRectangularGradient extends Com4jObject {
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
   * Getter method for the COM property "ColorStops"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ColorStops
   */

  @VTID(10)
  com.exceljava.com4j.excel.ColorStops getColorStops();


  /**
   * <p>
   * Getter method for the COM property "RectangleTop"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(11)
  double getRectangleTop();


  /**
   * <p>
   * Setter method for the COM property "RectangleTop"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(12)
  void setRectangleTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleBottom"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(13)
  double getRectangleBottom();


  /**
   * <p>
   * Setter method for the COM property "RectangleBottom"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(14)
  void setRectangleBottom(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleLeft"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(15)
  double getRectangleLeft();


  /**
   * <p>
   * Setter method for the COM property "RectangleLeft"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(16)
  void setRectangleLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleRight"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(17)
  double getRectangleRight();


  /**
   * <p>
   * Setter method for the COM property "RectangleRight"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(18)
  void setRectangleRight(
    double rhs);


  // Properties:
}
