package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface AutoRecover extends Com4jObject {
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
   * Getter method for the COM property "Time"
   * </p>
   */

  @DISPID(394)
  @PropGet
  int getTime();


  /**
   * <p>
   * Setter method for the COM property "Time"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(394)
  @PropPut
  void setTime(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Path"
   * </p>
   */

  @DISPID(291)
  @PropGet
  java.lang.String getPath();


  /**
   * <p>
   * Setter method for the COM property "Path"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(291)
  @PropPut
  void setPath(
    java.lang.String rhs);


  // Properties:
}
