package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface LinearGradient extends Com4jObject {
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
   * Getter method for the COM property "Degree"
   * </p>
   */

  @DISPID(1623)
  @PropGet
  double getDegree();


  /**
   * <p>
   * Setter method for the COM property "Degree"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1623)
  @PropPut
  void setDegree(
    double rhs);


  // Properties:
}
