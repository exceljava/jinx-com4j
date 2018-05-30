package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ValueChange extends Com4jObject {
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
   * Getter method for the COM property "Order"
   * </p>
   */

  @DISPID(192)
  @PropGet
  int getOrder();


  /**
   * <p>
   * Getter method for the COM property "VisibleInPivotTable"
   * </p>
   */

  @DISPID(2971)
  @PropGet
  boolean getVisibleInPivotTable();


  /**
   * <p>
   * Getter method for the COM property "PivotCell"
   * </p>
   */

  @DISPID(2013)
  @PropGet
  com.exceljava.com4j.excel.PivotCell getPivotCell();


  /**
   * <p>
   * Getter method for the COM property "Tuple"
   * </p>
   */

  @DISPID(2972)
  @PropGet
  java.lang.String getTuple();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  double getValue();


  /**
   * <p>
   * Getter method for the COM property "AllocationValue"
   * </p>
   */

  @DISPID(2874)
  @PropGet
  com.exceljava.com4j.excel.XlAllocationValue getAllocationValue();


  /**
   * <p>
   * Getter method for the COM property "AllocationMethod"
   * </p>
   */

  @DISPID(2875)
  @PropGet
  com.exceljava.com4j.excel.XlAllocationMethod getAllocationMethod();


  /**
   * <p>
   * Getter method for the COM property "AllocationWeightExpression"
   * </p>
   */

  @DISPID(2876)
  @PropGet
  java.lang.String getAllocationWeightExpression();


  /**
   */

  @DISPID(117)
  void delete();


  // Properties:
}
