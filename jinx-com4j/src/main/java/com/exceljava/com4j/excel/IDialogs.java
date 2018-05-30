package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020879-0001-0000-C000-000000000046}")
public interface IDialogs extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory com.exceljava.com4j.excel.XlBuiltInDialog parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Dialog
   */

  @VTID(11)
  com.exceljava.com4j.excel.Dialog getItem(
    com.exceljava.com4j.excel.XlBuiltInDialog index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory com.exceljava.com4j.excel.XlBuiltInDialog parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Dialog
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.Dialog get_Default(
    com.exceljava.com4j.excel.XlBuiltInDialog index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
