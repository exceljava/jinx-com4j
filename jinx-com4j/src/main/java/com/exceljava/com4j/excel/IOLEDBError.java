package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024445-0001-0000-C000-000000000046}")
public interface IOLEDBError extends Com4jObject {
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
   * Getter method for the COM property "SqlState"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getSqlState();


  /**
   * <p>
   * Getter method for the COM property "ErrorString"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getErrorString();


  /**
   * <p>
   * Getter method for the COM property "Native"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getNative();


  /**
   * <p>
   * Getter method for the COM property "Number"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getNumber();


  /**
   * <p>
   * Getter method for the COM property "Stage"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
  int getStage();


  // Properties:
}
