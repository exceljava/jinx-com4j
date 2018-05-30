package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface MultiThreadedCalculation extends Com4jObject {
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
   * Getter method for the COM property "Enabled"
   * </p>
   */

  @DISPID(600)
  @PropGet
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(600)
  @PropPut
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ThreadMode"
   * </p>
   */

  @DISPID(2766)
  @PropGet
  com.exceljava.com4j.excel.XlThreadMode getThreadMode();


  /**
   * <p>
   * Setter method for the COM property "ThreadMode"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlThreadMode parameter.
   */

  @DISPID(2766)
  @PropPut
  void setThreadMode(
    com.exceljava.com4j.excel.XlThreadMode rhs);


  /**
   * <p>
   * Getter method for the COM property "ThreadCount"
   * </p>
   */

  @DISPID(2767)
  @PropGet
  int getThreadCount();


  /**
   * <p>
   * Setter method for the COM property "ThreadCount"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2767)
  @PropPut
  void setThreadCount(
    int rhs);


  // Properties:
}
