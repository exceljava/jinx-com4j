package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C170D-0000-0000-C000-000000000046}")
public interface Points extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(8)
  int getCount();


  /**
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.ChartPoint
   */

  @VTID(9)
  com.exceljava.com4j.office.ChartPoint item(
    int index);


  /**
   */

  @VTID(10)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(11)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.ChartPoint
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.office.ChartPoint get_Default(
    int index);


  // Properties:
}
