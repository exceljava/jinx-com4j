package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024473-0001-0000-C000-000000000046}")
public interface IListColumn extends Com4jObject {
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
   */

  @VTID(10)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "ListDataFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListDataFormat
   */

  @VTID(12)
  com.exceljava.com4j.excel.ListDataFormat getListDataFormat();


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(15)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(16)
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "TotalsCalculation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTotalsCalculation
   */

  @VTID(17)
  com.exceljava.com4j.excel.XlTotalsCalculation getTotalsCalculation();


  /**
   * <p>
   * Setter method for the COM property "TotalsCalculation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTotalsCalculation parameter.
   */

  @VTID(18)
  void setTotalsCalculation(
    com.exceljava.com4j.excel.XlTotalsCalculation rhs);


  /**
   * <p>
   * Getter method for the COM property "XPath"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XPath
   */

  @VTID(19)
  com.exceljava.com4j.excel.XPath getXPath();


  /**
   * <p>
   * Getter method for the COM property "SharePointFormula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(20)
  java.lang.String getSharePointFormula();


  /**
   * <p>
   * Getter method for the COM property "DataBodyRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(21)
  com.exceljava.com4j.excel.Range getDataBodyRange();


  /**
   * <p>
   * Getter method for the COM property "Total"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(22)
  com.exceljava.com4j.excel.Range getTotal();


  // Properties:
}
