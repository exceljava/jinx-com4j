package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PublishedDoc extends Com4jObject {
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
   * Getter method for the COM property "Title"
   * </p>
   */

  @DISPID(199)
  @PropGet
  java.lang.String getTitle();


  /**
   * <p>
   * Getter method for the COM property "DisclosureScope"
   * </p>
   */

  @DISPID(3195)
  @PropGet
  int getDisclosureScope();


  /**
   * <p>
   * Getter method for the COM property "Url"
   * </p>
   */

  @DISPID(2271)
  @PropGet
  java.lang.String getUrl();


  /**
   * <p>
   * Getter method for the COM property "PublishedDate"
   * </p>
   */

  @DISPID(3230)
  @PropGet
  java.util.Date getPublishedDate();


  // Properties:
}
