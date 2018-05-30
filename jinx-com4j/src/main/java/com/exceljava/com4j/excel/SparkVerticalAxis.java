package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SparkVerticalAxis extends Com4jObject {
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
   * Getter method for the COM property "MinScaleType"
   * </p>
   */

  @DISPID(2965)
  @PropGet
  com.exceljava.com4j.excel.XlSparkScale getMinScaleType();


  /**
   * <p>
   * Setter method for the COM property "MinScaleType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSparkScale parameter.
   */

  @DISPID(2965)
  @PropPut
  void setMinScaleType(
    com.exceljava.com4j.excel.XlSparkScale rhs);


  /**
   * <p>
   * Getter method for the COM property "CustomMinScaleValue"
   * </p>
   */

  @DISPID(2966)
  @PropGet
  java.lang.Object getCustomMinScaleValue();


  /**
   * <p>
   * Setter method for the COM property "CustomMinScaleValue"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2966)
  @PropPut
  void setCustomMinScaleValue(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "MaxScaleType"
   * </p>
   */

  @DISPID(2967)
  @PropGet
  com.exceljava.com4j.excel.XlSparkScale getMaxScaleType();


  /**
   * <p>
   * Setter method for the COM property "MaxScaleType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSparkScale parameter.
   */

  @DISPID(2967)
  @PropPut
  void setMaxScaleType(
    com.exceljava.com4j.excel.XlSparkScale rhs);


  /**
   * <p>
   * Getter method for the COM property "CustomMaxScaleValue"
   * </p>
   */

  @DISPID(2968)
  @PropGet
  java.lang.Object getCustomMaxScaleValue();


  /**
   * <p>
   * Setter method for the COM property "CustomMaxScaleValue"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2968)
  @PropPut
  void setCustomMaxScaleValue(
    java.lang.Object rhs);


  // Properties:
}
