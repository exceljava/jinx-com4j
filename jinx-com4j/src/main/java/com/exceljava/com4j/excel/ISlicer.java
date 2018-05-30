package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C8-0001-0000-C000-000000000046}")
public interface ISlicer extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(11)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(13)
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(14)
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(15)
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(16)
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(17)
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "DisableMoveResizeUI"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getDisableMoveResizeUI();


  /**
   * <p>
   * Setter method for the COM property "DisableMoveResizeUI"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setDisableMoveResizeUI(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(20)
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(21)
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(22)
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(23)
  void setHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "RowHeight"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(24)
  double getRowHeight();


  /**
   * <p>
   * Setter method for the COM property "RowHeight"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(25)
  void setRowHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "ColumnWidth"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(26)
  double getColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "ColumnWidth"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(27)
  void setColumnWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberOfColumns"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(28)
  int getNumberOfColumns();


  /**
   * <p>
   * Setter method for the COM property "NumberOfColumns"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(29)
  void setNumberOfColumns(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayHeader"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getDisplayHeader();


  /**
   * <p>
   * Setter method for the COM property "DisplayHeader"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setDisplayHeader(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(32)
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(33)
  void setLocked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SlicerCache"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(34)
  com.exceljava.com4j.excel.SlicerCache getSlicerCache();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCacheLevel
   */

  @VTID(35)
  com.exceljava.com4j.excel.SlicerCacheLevel getSlicerCacheLevel();


  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(36)
  com.exceljava.com4j.excel.Shape getShape();


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(37)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(38)
  void setStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @VTID(39)
  void delete();


  /**
   */

  @VTID(40)
  void cut();


  /**
   */

  @VTID(41)
  void copy();


  /**
   * <p>
   * Getter method for the COM property "ActiveItem"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerItem
   */

  @VTID(42)
  com.exceljava.com4j.excel.SlicerItem getActiveItem();


  /**
   * <p>
   * Getter method for the COM property "TimelineViewState"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TimelineViewState
   */

  @VTID(43)
  com.exceljava.com4j.excel.TimelineViewState getTimelineViewState();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerCacheType
   */

  @VTID(44)
  com.exceljava.com4j.excel.XlSlicerCacheType getSlicerCacheType();


  // Properties:
}
