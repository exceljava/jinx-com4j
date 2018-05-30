package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface CustomView extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "PrintSettings"
   * </p>
   */

  @DISPID(1577)
  @PropGet
  boolean getPrintSettings();


  /**
   * <p>
   * Getter method for the COM property "RowColSettings"
   * </p>
   */

  @DISPID(1578)
  @PropGet
  boolean getRowColSettings();


  /**
   */

  @DISPID(496)
  void show();


  /**
   */

  @DISPID(117)
  void delete();


  // Properties:
}
