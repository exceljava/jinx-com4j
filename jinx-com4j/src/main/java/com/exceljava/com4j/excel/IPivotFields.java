package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020875-0001-0000-C000-000000000046}")
public interface IPivotFields extends Com4jObject,Iterable<Com4jObject> {
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @VTID(9)
  com.exceljava.com4j.excel.PivotTable getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(11)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
