package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Filter extends Com4jObject {
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
   * Getter method for the COM property "On"
   * </p>
   */

  @DISPID(1618)
  @PropGet
  boolean getOn();


  /**
   * <p>
   * Getter method for the COM property "Criteria1"
   * </p>
   */

  @DISPID(796)
  @PropGet
  java.lang.Object getCriteria1();


  /**
   * <p>
   * Getter method for the COM property "_Operator"
   * </p>
   */

  @DISPID(2641)
  @PropGet
  com.exceljava.com4j.excel.XlAutoFilterOperator get_Operator();


  /**
   * <p>
   * Getter method for the COM property "Criteria2"
   * </p>
   */

  @DISPID(798)
  @PropGet
  java.lang.Object getCriteria2();


  /**
   * <p>
   * Getter method for the COM property "Operator"
   * </p>
   */

  @DISPID(797)
  @PropGet
  com.exceljava.com4j.excel.XlAutoFilterOperator getOperator();


  /**
   * <p>
   * Setter method for the COM property "Operator"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlAutoFilterOperator parameter.
   */

  @DISPID(797)
  @PropPut
  void setOperator(
    com.exceljava.com4j.excel.XlAutoFilterOperator rhs);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  // Properties:
}
