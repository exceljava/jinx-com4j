package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244BA-0001-0000-C000-000000000046}")
public interface ISparkAxes extends Com4jObject {
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
   * Getter method for the COM property "Vertical"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkVerticalAxis
   */

  @VTID(10)
  com.exceljava.com4j.excel.SparkVerticalAxis getVertical();


  /**
   * <p>
   * Getter method for the COM property "Horizontal"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkHorizontalAxis
   */

  @VTID(11)
  com.exceljava.com4j.excel.SparkHorizontalAxis getHorizontal();


  // Properties:
}
