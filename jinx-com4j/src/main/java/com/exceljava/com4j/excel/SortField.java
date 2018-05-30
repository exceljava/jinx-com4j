package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SortField extends Com4jObject {
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
   * Getter method for the COM property "SortOn"
   * </p>
   */

  @DISPID(2741)
  @PropGet
  com.exceljava.com4j.excel.XlSortOn getSortOn();


  /**
   * <p>
   * Setter method for the COM property "SortOn"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortOn parameter.
   */

  @DISPID(2741)
  @PropPut
  void setSortOn(
    com.exceljava.com4j.excel.XlSortOn rhs);


  /**
   * <p>
   * Getter method for the COM property "SortOnValue"
   * </p>
   */

  @DISPID(2742)
  @PropGet
  com4j.Com4jObject getSortOnValue();


  /**
   * <p>
   * Getter method for the COM property "Key"
   * </p>
   */

  @DISPID(155)
  @PropGet
  com.exceljava.com4j.excel.Range getKey();


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   */

  @DISPID(192)
  @PropGet
  com.exceljava.com4j.excel.XlSortOrder getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortOrder parameter.
   */

  @DISPID(192)
  @PropPut
  void setOrder(
    com.exceljava.com4j.excel.XlSortOrder rhs);


  /**
   * <p>
   * Getter method for the COM property "CustomOrder"
   * </p>
   */

  @DISPID(2743)
  @PropGet
  java.lang.Object getCustomOrder();


  /**
   * <p>
   * Setter method for the COM property "CustomOrder"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2743)
  @PropPut
  void setCustomOrder(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DataOption"
   * </p>
   */

  @DISPID(2744)
  @PropGet
  com.exceljava.com4j.excel.XlSortDataOption getDataOption();


  /**
   * <p>
   * Setter method for the COM property "DataOption"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortDataOption parameter.
   */

  @DISPID(2744)
  @PropPut
  void setDataOption(
    com.exceljava.com4j.excel.XlSortDataOption rhs);


  /**
   * <p>
   * Getter method for the COM property "Priority"
   * </p>
   */

  @DISPID(985)
  @PropGet
  int getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(985)
  @PropPut
  void setPriority(
    int rhs);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(2745)
  void modifyKey(
    com.exceljava.com4j.excel.Range key);


  /**
   * @param icon Mandatory com.exceljava.com4j.excel.Icon parameter.
   */

  @DISPID(2746)
  void setIcon(
    com.exceljava.com4j.excel.Icon icon);


  // Properties:
}
