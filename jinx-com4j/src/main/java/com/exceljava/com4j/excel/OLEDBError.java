package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface OLEDBError extends Com4jObject {
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
   * Getter method for the COM property "SqlState"
   * </p>
   */

  @DISPID(1603)
  @PropGet
  java.lang.String getSqlState();


  /**
   * <p>
   * Getter method for the COM property "ErrorString"
   * </p>
   */

  @DISPID(1490)
  @PropGet
  java.lang.String getErrorString();


  /**
   * <p>
   * Getter method for the COM property "Native"
   * </p>
   */

  @DISPID(1897)
  @PropGet
  int getNative();


  /**
   * <p>
   * Getter method for the COM property "Number"
   * </p>
   */

  @DISPID(451)
  @PropGet
  int getNumber();


  /**
   * <p>
   * Getter method for the COM property "Stage"
   * </p>
   */

  @DISPID(1898)
  @PropGet
  int getStage();


  // Properties:
}
