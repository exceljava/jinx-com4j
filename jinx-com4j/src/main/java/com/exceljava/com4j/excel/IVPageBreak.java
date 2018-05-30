package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024402-0001-0000-C000-000000000046}")
public interface IVPageBreak extends Com4jObject {
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
   * @return  Returns a value of type com.exceljava.com4j.excel._Worksheet
   */

  @VTID(9)
  com.exceljava.com4j.excel._Worksheet getParent();


  /**
   */

  @VTID(10)
  void delete();


  /**
   * @param direction Mandatory com.exceljava.com4j.excel.XlDirection parameter.
   * @param regionIndex Mandatory int parameter.
   */

  @VTID(11)
  void dragOff(
    com.exceljava.com4j.excel.XlDirection direction,
    int regionIndex);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPageBreak
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlPageBreak getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPageBreak parameter.
   */

  @VTID(13)
  void setType(
    com.exceljava.com4j.excel.XlPageBreak rhs);


  /**
   * <p>
   * Getter method for the COM property "Extent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPageBreakExtent
   */

  @VTID(14)
  com.exceljava.com4j.excel.XlPageBreakExtent getExtent();


  /**
   * <p>
   * Getter method for the COM property "Location"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(15)
  com.exceljava.com4j.excel.Range getLocation();


  /**
   * <p>
   * Setter method for the COM property "Location"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(16)
  void setLocation(
    com.exceljava.com4j.excel.Range rhs);


  // Properties:
}
