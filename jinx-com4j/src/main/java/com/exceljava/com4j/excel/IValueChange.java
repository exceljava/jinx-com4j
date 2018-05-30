package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C0-0001-0000-C000-000000000046}")
public interface IValueChange extends Com4jObject {
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
   * Getter method for the COM property "VisibleInPivotTable"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getVisibleInPivotTable();


  /**
   * <p>
   * Getter method for the COM property "PivotCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCell
   */

  @VTID(12)
  com.exceljava.com4j.excel.PivotCell getPivotCell();


  /**
   * <p>
   * Getter method for the COM property "Tuple"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getTuple();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(14)
  double getValue();


  /**
   * <p>
   * Getter method for the COM property "AllocationValue"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlAllocationValue
   */

  @VTID(15)
  com.exceljava.com4j.excel.XlAllocationValue getAllocationValue();


  /**
   * <p>
   * Getter method for the COM property "AllocationMethod"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlAllocationMethod
   */

  @VTID(16)
  com.exceljava.com4j.excel.XlAllocationMethod getAllocationMethod();


  /**
   * <p>
   * Getter method for the COM property "AllocationWeightExpression"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(17)
  java.lang.String getAllocationWeightExpression();


  /**
   */

  @VTID(18)
  void delete();


  // Properties:
}
