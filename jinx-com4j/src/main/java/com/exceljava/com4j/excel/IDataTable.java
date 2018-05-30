package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020843-0001-0000-C000-000000000046}")
public interface IDataTable extends Com4jObject {
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
   * Getter method for the COM property "ShowLegendKey"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getShowLegendKey();


  /**
   * <p>
   * Setter method for the COM property "ShowLegendKey"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setShowLegendKey(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasBorderHorizontal"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getHasBorderHorizontal();


  /**
   * <p>
   * Setter method for the COM property "HasBorderHorizontal"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(13)
  void setHasBorderHorizontal(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasBorderVertical"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getHasBorderVertical();


  /**
   * <p>
   * Setter method for the COM property "HasBorderVertical"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setHasBorderVertical(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasBorderOutline"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getHasBorderOutline();


  /**
   * <p>
   * Setter method for the COM property "HasBorderOutline"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setHasBorderOutline(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Border
   */

  @VTID(18)
  com.exceljava.com4j.excel.Border getBorder();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(19)
  com.exceljava.com4j.excel.Font getFont();


  /**
   */

  @VTID(20)
  void select();


  /**
   */

  @VTID(21)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "AutoScaleFont"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAutoScaleFont();


  /**
   * <p>
   * Setter method for the COM property "AutoScaleFont"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(23)
  void setAutoScaleFont(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartFormat
   */

  @VTID(24)
  com.exceljava.com4j.excel.ChartFormat getFormat();


  // Properties:
}
