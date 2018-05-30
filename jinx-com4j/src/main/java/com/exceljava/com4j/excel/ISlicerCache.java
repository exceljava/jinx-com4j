package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C4-0001-0000-C000-000000000046}")
public interface ISlicerCache extends Com4jObject {
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
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "OLAP"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getOLAP();


  /**
   * <p>
   * Getter method for the COM property "SourceType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotTableSourceType
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlPivotTableSourceType getSourceType();


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(13)
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  /**
   * <p>
   * Getter method for the COM property "Slicers"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Slicers
   */

  @VTID(14)
  com.exceljava.com4j.excel.Slicers getSlicers();


  /**
   * <p>
   * Getter method for the COM property "PivotTables"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerPivotTables
   */

  @VTID(15)
  com.exceljava.com4j.excel.SlicerPivotTables getPivotTables();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheLevels"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCacheLevels
   */

  @VTID(16)
  com.exceljava.com4j.excel.SlicerCacheLevels getSlicerCacheLevels();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(17)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(18)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleSlicerItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerItems
   */

  @VTID(19)
  com.exceljava.com4j.excel.SlicerItems getVisibleSlicerItems();


  /**
   * <p>
   * Getter method for the COM property "VisibleSlicerItemsList"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(20)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVisibleSlicerItemsList();


  /**
   * <p>
   * Setter method for the COM property "VisibleSlicerItemsList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(21)
  void setVisibleSlicerItemsList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SlicerItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerItems
   */

  @VTID(22)
  com.exceljava.com4j.excel.SlicerItems getSlicerItems();


  /**
   * <p>
   * Getter method for the COM property "CrossFilterType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerCrossFilterType
   */

  @VTID(23)
  com.exceljava.com4j.excel.XlSlicerCrossFilterType getCrossFilterType();


  /**
   * <p>
   * Setter method for the COM property "CrossFilterType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerCrossFilterType parameter.
   */

  @VTID(24)
  void setCrossFilterType(
    com.exceljava.com4j.excel.XlSlicerCrossFilterType rhs);


  /**
   * <p>
   * Getter method for the COM property "SortItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerSort
   */

  @VTID(25)
  com.exceljava.com4j.excel.XlSlicerSort getSortItems();


  /**
   * <p>
   * Setter method for the COM property "SortItems"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerSort parameter.
   */

  @VTID(26)
  void setSortItems(
    com.exceljava.com4j.excel.XlSlicerSort rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "SortUsingCustomLists"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getSortUsingCustomLists();


  /**
   * <p>
   * Setter method for the COM property "SortUsingCustomLists"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setSortUsingCustomLists(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAllItems"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getShowAllItems();


  /**
   * <p>
   * Setter method for the COM property "ShowAllItems"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setShowAllItems(
    boolean rhs);


  /**
   */

  @VTID(32)
  void clearManualFilter();


  /**
   */

  @VTID(33)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "TimelineState"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TimelineState
   */

  @VTID(34)
  com.exceljava.com4j.excel.TimelineState getTimelineState();


  /**
   */

  @VTID(35)
  void clearAllFilters();


  /**
   * <p>
   * Getter method for the COM property "SlicerCacheType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerCacheType
   */

  @VTID(36)
  com.exceljava.com4j.excel.XlSlicerCacheType getSlicerCacheType();


  /**
   * <p>
   * Getter method for the COM property "FilterCleared"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(37)
  boolean getFilterCleared();


  /**
   * <p>
   * Getter method for the COM property "List"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(38)
  boolean getList();


  /**
   * <p>
   * Getter method for the COM property "RequireManualUpdate"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(39)
  boolean getRequireManualUpdate();


  /**
   * <p>
   * Setter method for the COM property "RequireManualUpdate"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(40)
  void setRequireManualUpdate(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(41)
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   */

  @VTID(42)
  void clearDateFilter();


  // Properties:
}
