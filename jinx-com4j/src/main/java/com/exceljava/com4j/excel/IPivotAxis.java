package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024482-0001-0000-C000-000000000046}")
public interface IPivotAxis extends Com4jObject {
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
   * Getter method for the COM property "PivotLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotLines
   */

  @VTID(10)
  com.exceljava.com4j.excel.PivotLines getPivotLines();


  // Properties:
}
