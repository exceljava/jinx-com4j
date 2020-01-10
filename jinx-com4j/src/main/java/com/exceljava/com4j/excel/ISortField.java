package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244A9-0001-0000-C000-000000000046}")
public interface ISortField extends Com4jObject {
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
   * Getter method for the COM property "SortOn"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSortOn
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlSortOn getSortOn();


  /**
   * <p>
   * Setter method for the COM property "SortOn"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortOn parameter.
   */

  @VTID(11)
  void setSortOn(
    com.exceljava.com4j.excel.XlSortOn rhs);


  /**
   * <p>
   * Getter method for the COM property "SortOnValue"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getSortOnValue();


  /**
   * <p>
   * Getter method for the COM property "Key"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(13)
  com.exceljava.com4j.excel.Range getKey();


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSortOrder
   */

  @VTID(14)
  com.exceljava.com4j.excel.XlSortOrder getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortOrder parameter.
   */

  @VTID(15)
  void setOrder(
    com.exceljava.com4j.excel.XlSortOrder rhs);


  /**
   * <p>
   * Getter method for the COM property "CustomOrder"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCustomOrder();


  /**
   * <p>
   * Setter method for the COM property "CustomOrder"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(17)
  void setCustomOrder(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DataOption"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSortDataOption
   */

  @VTID(18)
  com.exceljava.com4j.excel.XlSortDataOption getDataOption();


  /**
   * <p>
   * Setter method for the COM property "DataOption"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortDataOption parameter.
   */

  @VTID(19)
  void setDataOption(
    com.exceljava.com4j.excel.XlSortDataOption rhs);


  /**
   * <p>
   * Getter method for the COM property "Priority"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(20)
  int getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(21)
  void setPriority(
    int rhs);


  /**
   */

  @VTID(22)
  void delete();


  /**
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(23)
  void modifyKey(
    com.exceljava.com4j.excel.Range key);


  /**
   * @param icon Mandatory com.exceljava.com4j.excel.Icon parameter.
   */

  @VTID(24)
  void setIcon(
    com.exceljava.com4j.excel.Icon icon);


  /**
   * <p>
   * Getter method for the COM property "SubField"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(25)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSubField();


  /**
   * <p>
   * Setter method for the COM property "SubField"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(26)
  void setSubField(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  // Properties:
}
