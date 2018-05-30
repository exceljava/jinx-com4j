package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024406-0001-0000-C000-000000000046}")
public interface IRecentFiles extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Maximum"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getMaximum();


  /**
   * <p>
   * Setter method for the COM property "Maximum"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(11)
  void setMaximum(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.RecentFile
   */

  @VTID(13)
  com.exceljava.com4j.excel.RecentFile getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.RecentFile
   */

  @VTID(14)
  @DefaultMethod
  com.exceljava.com4j.excel.RecentFile get_Default(
    int index);


  /**
   * @param name Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.RecentFile
   */

  @VTID(15)
  com.exceljava.com4j.excel.RecentFile add(
    java.lang.String name);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(16)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
