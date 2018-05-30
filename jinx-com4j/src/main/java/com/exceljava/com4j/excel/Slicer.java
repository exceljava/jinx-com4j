package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Slicer extends Com4jObject {
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
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   */

  @DISPID(139)
  @PropGet
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(139)
  @PropPut
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(126)
  @PropPut
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(127)
  @PropPut
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "DisableMoveResizeUI"
   * </p>
   */

  @DISPID(2983)
  @PropGet
  boolean getDisableMoveResizeUI();


  /**
   * <p>
   * Setter method for the COM property "DisableMoveResizeUI"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2983)
  @PropPut
  void setDisableMoveResizeUI(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RowHeight"
   * </p>
   */

  @DISPID(272)
  @PropGet
  double getRowHeight();


  /**
   * <p>
   * Setter method for the COM property "RowHeight"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(272)
  @PropPut
  void setRowHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "ColumnWidth"
   * </p>
   */

  @DISPID(242)
  @PropGet
  double getColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "ColumnWidth"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(242)
  @PropPut
  void setColumnWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberOfColumns"
   * </p>
   */

  @DISPID(2984)
  @PropGet
  int getNumberOfColumns();


  /**
   * <p>
   * Setter method for the COM property "NumberOfColumns"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2984)
  @PropPut
  void setNumberOfColumns(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayHeader"
   * </p>
   */

  @DISPID(2985)
  @PropGet
  boolean getDisplayHeader();


  /**
   * <p>
   * Setter method for the COM property "DisplayHeader"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2985)
  @PropPut
  void setDisplayHeader(
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
   * <p>
   * Getter method for the COM property "SlicerCache"
   * </p>
   */

  @DISPID(2986)
  @PropGet
  com.exceljava.com4j.excel.SlicerCache getSlicerCache();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheLevel"
   * </p>
   */

  @DISPID(2987)
  @PropGet
  com.exceljava.com4j.excel.SlicerCacheLevel getSlicerCacheLevel();


  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   */

  @DISPID(1582)
  @PropGet
  com.exceljava.com4j.excel.Shape getShape();


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   */

  @DISPID(260)
  @PropGet
  java.lang.Object getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(260)
  @PropPut
  void setStyle(
    java.lang.Object rhs);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   */

  @DISPID(565)
  void cut();


  /**
   */

  @DISPID(551)
  void copy();


  /**
   * <p>
   * Getter method for the COM property "ActiveItem"
   * </p>
   */

  @DISPID(2988)
  @PropGet
  com.exceljava.com4j.excel.SlicerItem getActiveItem();


  /**
   * <p>
   * Getter method for the COM property "TimelineViewState"
   * </p>
   */

  @DISPID(3116)
  @PropGet
  com.exceljava.com4j.excel.TimelineViewState getTimelineViewState();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheType"
   * </p>
   */

  @DISPID(3111)
  @PropGet
  com.exceljava.com4j.excel.XlSlicerCacheType getSlicerCacheType();


  // Properties:
}
