package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024494-0001-0000-C000-000000000046}")
public interface IColorScaleCriteria extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(7)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ColorScaleCriterion
   */

  @VTID(8)
  @DefaultMethod
  com.exceljava.com4j.excel.ColorScaleCriterion get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(9)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ColorScaleCriterion
   */

  @VTID(10)
  com.exceljava.com4j.excel.ColorScaleCriterion getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  // Properties:
}
