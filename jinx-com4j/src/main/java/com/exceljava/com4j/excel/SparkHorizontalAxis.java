package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SparkHorizontalAxis extends Com4jObject {
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
   * Getter method for the COM property "Axis"
   * </p>
   */

  @DISPID(67)
  @PropGet
  com.exceljava.com4j.excel.SparkColor getAxis();


  /**
   * <p>
   * Getter method for the COM property "IsDateAxis"
   * </p>
   */

  @DISPID(2963)
  @PropGet
  boolean getIsDateAxis();


  /**
   * <p>
   * Getter method for the COM property "RightToLeftPlotOrder"
   * </p>
   */

  @DISPID(2964)
  @PropGet
  boolean getRightToLeftPlotOrder();


  /**
   * <p>
   * Setter method for the COM property "RightToLeftPlotOrder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2964)
  @PropPut
  void setRightToLeftPlotOrder(
    boolean rhs);


  // Properties:
}
