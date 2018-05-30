package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface RectangularGradient extends Com4jObject {
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
   * Getter method for the COM property "ColorStops"
   * </p>
   */

  @DISPID(2761)
  @PropGet
  com.exceljava.com4j.excel.ColorStops getColorStops();


  /**
   * <p>
   * Getter method for the COM property "RectangleTop"
   * </p>
   */

  @DISPID(2762)
  @PropGet
  double getRectangleTop();


  /**
   * <p>
   * Setter method for the COM property "RectangleTop"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2762)
  @PropPut
  void setRectangleTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleBottom"
   * </p>
   */

  @DISPID(2763)
  @PropGet
  double getRectangleBottom();


  /**
   * <p>
   * Setter method for the COM property "RectangleBottom"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2763)
  @PropPut
  void setRectangleBottom(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleLeft"
   * </p>
   */

  @DISPID(2764)
  @PropGet
  double getRectangleLeft();


  /**
   * <p>
   * Setter method for the COM property "RectangleLeft"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2764)
  @PropPut
  void setRectangleLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RectangleRight"
   * </p>
   */

  @DISPID(2765)
  @PropGet
  double getRectangleRight();


  /**
   * <p>
   * Setter method for the COM property "RectangleRight"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2765)
  @PropPut
  void setRectangleRight(
    double rhs);


  // Properties:
}
