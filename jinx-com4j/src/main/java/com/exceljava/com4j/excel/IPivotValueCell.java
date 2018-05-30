package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244CF-0001-0000-C000-000000000046}")
public interface IPivotValueCell extends Com4jObject {
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
   * Getter method for the COM property "PivotCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCell
   */

  @VTID(10)
  com.exceljava.com4j.excel.PivotCell getPivotCell();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue();


  /**
   */

  @VTID(12)
  void showDetail();


  /**
   * <p>
   * Getter method for the COM property "ServerActions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Actions
   */

  @VTID(13)
  com.exceljava.com4j.excel.Actions getServerActions();


  // Properties:
}
