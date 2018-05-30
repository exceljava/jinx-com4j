package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelFormatScientificNumber extends Com4jObject {
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
   * Getter method for the COM property "DecimalPlaces"
   * </p>
   */

  @DISPID(2349)
  @PropGet
  int getDecimalPlaces();


  /**
   * <p>
   * Setter method for the COM property "DecimalPlaces"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2349)
  @PropPut
  void setDecimalPlaces(
    int rhs);


  // Properties:
}
