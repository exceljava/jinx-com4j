package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface LinkFormat extends Com4jObject {
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
   * Getter method for the COM property "AutoUpdate"
   * </p>
   */

  @DISPID(1048)
  @PropGet
  boolean getAutoUpdate();


  /**
   * <p>
   * Setter method for the COM property "AutoUpdate"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1048)
  @PropPut
  void setAutoUpdate(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   */

  @DISPID(269)
  @PropGet
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(269)
  @PropPut
  void setLocked(
    boolean rhs);


  /**
   */

  @DISPID(680)
  void update();


  // Properties:
}
