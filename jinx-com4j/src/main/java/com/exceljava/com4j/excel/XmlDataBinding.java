package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface XmlDataBinding extends Com4jObject {
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
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.String get_Default();


  /**
   */

  @DISPID(1417)
  com.exceljava.com4j.excel.XlXmlImportResult refresh();


  /**
   * @param url Mandatory java.lang.String parameter.
   */

  @DISPID(2329)
  void loadSettings(
    java.lang.String url);


  /**
   */

  @DISPID(2330)
  void clearSettings();


  /**
   * <p>
   * Getter method for the COM property "SourceUrl"
   * </p>
   */

  @DISPID(2331)
  @PropGet
  java.lang.String getSourceUrl();


  // Properties:
}
