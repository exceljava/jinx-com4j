package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024405-0001-0000-C000-000000000046}")
public interface IVPageBreaks extends Com4jObject,Iterable<Com4jObject> {
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
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.VPageBreak
   */

  @VTID(11)
  com.exceljava.com4j.excel.VPageBreak getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.VPageBreak
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.VPageBreak get_Default(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param before Mandatory com4j.Com4jObject parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.VPageBreak
   */

  @VTID(14)
  com.exceljava.com4j.excel.VPageBreak add(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject before);


  // Properties:
}
