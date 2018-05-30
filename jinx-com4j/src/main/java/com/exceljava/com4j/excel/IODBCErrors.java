package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002442D-0001-0000-C000-000000000046}")
public interface IODBCErrors extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ODBCError
   */

  @VTID(11)
  com.exceljava.com4j.excel.ODBCError item(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ODBCError
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.ODBCError get_Default(
    int index);


  /**
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
