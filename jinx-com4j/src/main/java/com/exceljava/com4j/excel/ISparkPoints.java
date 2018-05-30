package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244B8-0001-0000-C000-000000000046}")
public interface ISparkPoints extends Com4jObject {
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
   * Getter method for the COM property "Negative"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(10)
  com.exceljava.com4j.excel.SparkColor getNegative();


  /**
   * <p>
   * Getter method for the COM property "Markers"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(11)
  com.exceljava.com4j.excel.SparkColor getMarkers();


  /**
   * <p>
   * Getter method for the COM property "Highpoint"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(12)
  com.exceljava.com4j.excel.SparkColor getHighpoint();


  /**
   * <p>
   * Getter method for the COM property "Lowpoint"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(13)
  com.exceljava.com4j.excel.SparkColor getLowpoint();


  /**
   * <p>
   * Getter method for the COM property "Firstpoint"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(14)
  com.exceljava.com4j.excel.SparkColor getFirstpoint();


  /**
   * <p>
   * Getter method for the COM property "Lastpoint"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkColor
   */

  @VTID(15)
  com.exceljava.com4j.excel.SparkColor getLastpoint();


  // Properties:
}
