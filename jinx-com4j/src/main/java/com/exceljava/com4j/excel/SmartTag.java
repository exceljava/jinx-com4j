package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SmartTag extends Com4jObject {
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
   * Getter method for the COM property "DownloadURL"
   * </p>
   */

  @DISPID(2212)
  @PropGet
  java.lang.String getDownloadURL();


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
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "XML"
   * </p>
   */

  @DISPID(2213)
  @PropGet
  java.lang.String getXML();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   */

  @DISPID(197)
  @PropGet
  com.exceljava.com4j.excel.Range getRange();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "SmartTagActions"
   * </p>
   */

  @DISPID(2214)
  @PropGet
  com.exceljava.com4j.excel.SmartTagActions getSmartTagActions();


  /**
   * <p>
   * Getter method for the COM property "Properties"
   * </p>
   */

  @DISPID(2135)
  @PropGet
  com.exceljava.com4j.excel.CustomProperties getProperties();


  // Properties:
}
