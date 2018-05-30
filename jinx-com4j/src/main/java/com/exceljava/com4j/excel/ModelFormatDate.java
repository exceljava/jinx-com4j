package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelFormatDate extends Com4jObject {
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
   * Getter method for the COM property "FormatString"
   * </p>
   */

  @DISPID(3213)
  @PropGet
  java.lang.String getFormatString();


  /**
   * <p>
   * Setter method for the COM property "FormatString"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(3213)
  @PropPut
  void setFormatString(
    java.lang.String rhs);


  // Properties:
}
