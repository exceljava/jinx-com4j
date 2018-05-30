package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244DF-0001-0000-C000-000000000046}")
public interface ITimelineState extends Com4jObject {
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
   * Getter method for the COM property "StartDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getStartDate();


  /**
   * <p>
   * Getter method for the COM property "EndDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getEndDate();


  /**
   * <p>
   * Getter method for the COM property "FilterType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotFilterType
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlPivotFilterType getFilterType();


  /**
   * <p>
   * Getter method for the COM property "FilterValue1"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFilterValue1();


  /**
   * <p>
   * Getter method for the COM property "FilterValue2"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFilterValue2();


  /**
   * <p>
   * Getter method for the COM property "SingleRangeFilterState"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getSingleRangeFilterState();


  /**
   * @param startDate Mandatory java.lang.Object parameter.
   * @param endDate Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlFilterStatus
   */

  @VTID(16)
  com.exceljava.com4j.excel.XlFilterStatus setFilterDateRange(
    @MarshalAs(NativeType.VARIANT) java.lang.Object startDate,
    @MarshalAs(NativeType.VARIANT) java.lang.Object endDate);


  // Properties:
}
