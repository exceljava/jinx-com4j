package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface IconCriterion extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   */

  @DISPID(486)
  @PropGet
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.excel.XlConditionValueTypes getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlConditionValueTypes parameter.
   */

  @DISPID(108)
  @PropPut
  void setType(
    com.exceljava.com4j.excel.XlConditionValueTypes rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.Object getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Operator"
   * </p>
   */

  @DISPID(797)
  @PropGet
  int getOperator();


  /**
   * <p>
   * Setter method for the COM property "Operator"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(797)
  @PropPut
  void setOperator(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Icon"
   * </p>
   */

  @DISPID(2747)
  @PropGet
  com.exceljava.com4j.excel.XlIcon getIcon();


  /**
   * <p>
   * Setter method for the COM property "Icon"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlIcon parameter.
   */

  @DISPID(2747)
  @PropPut
  void setIcon(
    com.exceljava.com4j.excel.XlIcon rhs);


  // Properties:
}
