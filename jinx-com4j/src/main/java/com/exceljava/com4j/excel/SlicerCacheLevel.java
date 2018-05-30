package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SlicerCacheLevel extends Com4jObject {
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
   * Getter method for the COM property "SlicerItems"
   * </p>
   */

  @DISPID(2977)
  @PropGet
  com.exceljava.com4j.excel.SlicerItems getSlicerItems();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Ordinal"
   * </p>
   */

  @DISPID(2981)
  @PropGet
  int getOrdinal();


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
   * Getter method for the COM property "CrossFilterType"
   * </p>
   */

  @DISPID(2978)
  @PropGet
  com.exceljava.com4j.excel.XlSlicerCrossFilterType getCrossFilterType();


  /**
   * <p>
   * Setter method for the COM property "CrossFilterType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerCrossFilterType parameter.
   */

  @DISPID(2978)
  @PropPut
  void setCrossFilterType(
    com.exceljava.com4j.excel.XlSlicerCrossFilterType rhs);


  /**
   * <p>
   * Getter method for the COM property "SortItems"
   * </p>
   */

  @DISPID(2979)
  @PropGet
  com.exceljava.com4j.excel.XlSlicerSort getSortItems();


  /**
   * <p>
   * Setter method for the COM property "SortItems"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerSort parameter.
   */

  @DISPID(2979)
  @PropPut
  void setSortItems(
    com.exceljava.com4j.excel.XlSlicerSort rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleSlicerItemsList"
   * </p>
   */

  @DISPID(2976)
  @PropGet
  java.lang.Object getVisibleSlicerItemsList();


  // Properties:
}
