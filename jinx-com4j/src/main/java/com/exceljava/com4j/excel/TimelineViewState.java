package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface TimelineViewState extends Com4jObject {
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
   * Getter method for the COM property "ShowHeader"
   * </p>
   */

  @DISPID(3139)
  @PropGet
  boolean getShowHeader();


  /**
   * <p>
   * Setter method for the COM property "ShowHeader"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3139)
  @PropPut
  void setShowHeader(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowSelectionLabel"
   * </p>
   */

  @DISPID(3140)
  @PropGet
  boolean getShowSelectionLabel();


  /**
   * <p>
   * Setter method for the COM property "ShowSelectionLabel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3140)
  @PropPut
  void setShowSelectionLabel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTimeLevel"
   * </p>
   */

  @DISPID(3141)
  @PropGet
  boolean getShowTimeLevel();


  /**
   * <p>
   * Setter method for the COM property "ShowTimeLevel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3141)
  @PropPut
  void setShowTimeLevel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowHorizontalScrollbar"
   * </p>
   */

  @DISPID(3142)
  @PropGet
  boolean getShowHorizontalScrollbar();


  /**
   * <p>
   * Setter method for the COM property "ShowHorizontalScrollbar"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3142)
  @PropPut
  void setShowHorizontalScrollbar(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Level"
   * </p>
   */

  @DISPID(2980)
  @PropGet
  com.exceljava.com4j.excel.XlTimelineLevel getLevel();


  /**
   * <p>
   * Setter method for the COM property "Level"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTimelineLevel parameter.
   */

  @DISPID(2980)
  @PropPut
  void setLevel(
    com.exceljava.com4j.excel.XlTimelineLevel rhs);


  // Properties:
}
