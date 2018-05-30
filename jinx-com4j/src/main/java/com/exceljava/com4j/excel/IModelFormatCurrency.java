package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244F5-0001-0000-C000-000000000046}")
public interface IModelFormatCurrency extends Com4jObject {
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
   * Getter method for the COM property "Symbol"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getSymbol();


  /**
   * <p>
   * Setter method for the COM property "Symbol"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(11)
  void setSymbol(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DecimalPlaces"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getDecimalPlaces();


  /**
   * <p>
   * Setter method for the COM property "DecimalPlaces"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(13)
  void setDecimalPlaces(
    int rhs);


  // Properties:
}
