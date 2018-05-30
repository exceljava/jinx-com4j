package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244A3-0001-0000-C000-000000000046}")
public interface IPages extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Page
   */

  @VTID(7)
  com.exceljava.com4j.excel.Page getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Page
   */

  @VTID(8)
  @DefaultMethod
  com.exceljava.com4j.excel.Page get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(9)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(10)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
