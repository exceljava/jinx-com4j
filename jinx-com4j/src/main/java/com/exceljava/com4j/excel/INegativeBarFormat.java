package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244BF-0001-0000-C000-000000000046}")
public interface INegativeBarFormat extends Com4jObject {
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
   * Getter method for the COM property "ColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDataBarNegativeColorType
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlDataBarNegativeColorType getColorType();


  /**
   * <p>
   * Setter method for the COM property "ColorType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDataBarNegativeColorType parameter.
   */

  @VTID(11)
  void setColorType(
    com.exceljava.com4j.excel.XlDataBarNegativeColorType rhs);


  /**
   * <p>
   * Getter method for the COM property "BorderColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDataBarNegativeColorType
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlDataBarNegativeColorType getBorderColorType();


  /**
   * <p>
   * Setter method for the COM property "BorderColorType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDataBarNegativeColorType parameter.
   */

  @VTID(13)
  void setBorderColorType(
    com.exceljava.com4j.excel.XlDataBarNegativeColorType rhs);


  /**
   * <p>
   * Getter method for the COM property "Color"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getColor();


  /**
   * <p>
   * Getter method for the COM property "BorderColor"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(15)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getBorderColor();


  // Properties:
}
