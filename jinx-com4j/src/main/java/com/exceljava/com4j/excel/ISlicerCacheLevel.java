package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C6-0001-0000-C000-000000000046}")
public interface ISlicerCacheLevel extends Com4jObject {
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
   * Getter method for the COM property "SlicerItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerItems
   */

  @VTID(10)
  com.exceljava.com4j.excel.SlicerItems getSlicerItems();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Ordinal"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getOrdinal();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "CrossFilterType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerCrossFilterType
   */

  @VTID(14)
  com.exceljava.com4j.excel.XlSlicerCrossFilterType getCrossFilterType();


  /**
   * <p>
   * Setter method for the COM property "CrossFilterType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerCrossFilterType parameter.
   */

  @VTID(15)
  void setCrossFilterType(
    com.exceljava.com4j.excel.XlSlicerCrossFilterType rhs);


  /**
   * <p>
   * Getter method for the COM property "SortItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSlicerSort
   */

  @VTID(16)
  com.exceljava.com4j.excel.XlSlicerSort getSortItems();


  /**
   * <p>
   * Setter method for the COM property "SortItems"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSlicerSort parameter.
   */

  @VTID(17)
  void setSortItems(
    com.exceljava.com4j.excel.XlSlicerSort rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleSlicerItemsList"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVisibleSlicerItemsList();


  // Properties:
}
