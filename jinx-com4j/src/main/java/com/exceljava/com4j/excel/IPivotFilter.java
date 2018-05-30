package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024483-0001-0000-C000-000000000046}")
public interface IPivotFilter extends Com4jObject {
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
   * Getter method for the COM property "Order"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(11)
  void setOrder(
    int rhs);


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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getDescription();


  /**
   */

  @VTID(15)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getActive();


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(17)
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "DataField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(18)
  com.exceljava.com4j.excel.PivotField getDataField();


  /**
   * <p>
   * Getter method for the COM property "DataCubeField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @VTID(19)
  com.exceljava.com4j.excel.CubeField getDataCubeField();


  /**
   * <p>
   * Getter method for the COM property "Value1"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(20)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue1();


  /**
   * <p>
   * Getter method for the COM property "Value2"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue2();


  /**
   * <p>
   * Getter method for the COM property "MemberPropertyField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(22)
  com.exceljava.com4j.excel.PivotField getMemberPropertyField();


  /**
   * <p>
   * Getter method for the COM property "IsMemberPropertyFilter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(23)
  boolean getIsMemberPropertyFilter();


  /**
   * <p>
   * Getter method for the COM property "WholeDayFilter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getWholeDayFilter();


  /**
   * <p>
   * Setter method for the COM property "WholeDayFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setWholeDayFilter(
    boolean rhs);


  // Properties:
}
