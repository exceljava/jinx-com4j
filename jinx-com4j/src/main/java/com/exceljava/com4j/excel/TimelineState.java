package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface TimelineState extends Com4jObject {
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
   * Getter method for the COM property "StartDate"
   * </p>
   */

  @DISPID(3133)
  @PropGet
  java.lang.Object getStartDate();


  /**
   * <p>
   * Getter method for the COM property "EndDate"
   * </p>
   */

  @DISPID(3134)
  @PropGet
  java.lang.Object getEndDate();


  /**
   * <p>
   * Getter method for the COM property "FilterType"
   * </p>
   */

  @DISPID(2686)
  @PropGet
  com.exceljava.com4j.excel.XlPivotFilterType getFilterType();


  /**
   * <p>
   * Getter method for the COM property "FilterValue1"
   * </p>
   */

  @DISPID(3135)
  @PropGet
  java.lang.Object getFilterValue1();


  /**
   * <p>
   * Getter method for the COM property "FilterValue2"
   * </p>
   */

  @DISPID(3136)
  @PropGet
  java.lang.Object getFilterValue2();


  /**
   * <p>
   * Getter method for the COM property "SingleRangeFilterState"
   * </p>
   */

  @DISPID(3137)
  @PropGet
  boolean getSingleRangeFilterState();


  /**
   * @param startDate Mandatory java.lang.Object parameter.
   * @param endDate Mandatory java.lang.Object parameter.
   */

  @DISPID(3138)
  com.exceljava.com4j.excel.XlFilterStatus setFilterDateRange(
    java.lang.Object startDate,
    java.lang.Object endDate);


  // Properties:
}
