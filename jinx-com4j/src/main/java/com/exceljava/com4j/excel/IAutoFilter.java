package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024432-0001-0000-C000-000000000046}")
public interface IAutoFilter extends Com4jObject {
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
   * Getter method for the COM property "Range"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(10)
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "Filters"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Filters
   */

  @VTID(11)
  com.exceljava.com4j.excel.Filters getFilters();


  /**
   * <p>
   * Getter method for the COM property "FilterMode"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getFilterMode();


  /**
   * <p>
   * Getter method for the COM property "Sort"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sort
   */

  @VTID(13)
  com.exceljava.com4j.excel.Sort getSort();


  /**
   */

  @VTID(14)
  void applyFilter();


  /**
   */

  @VTID(15)
  void showAllData();


  // Properties:
}
