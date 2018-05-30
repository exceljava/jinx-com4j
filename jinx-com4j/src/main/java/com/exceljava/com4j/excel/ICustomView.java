package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024423-0001-0000-C000-000000000046}")
public interface ICustomView extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "PrintSettings"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getPrintSettings();


  /**
   * <p>
   * Getter method for the COM property "RowColSettings"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getRowColSettings();


  /**
   */

  @VTID(13)
  void show();


  /**
   */

  @VTID(14)
  void delete();


  // Properties:
}
