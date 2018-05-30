package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208C9-0001-0000-C000-000000000046}")
public interface ITickLabels extends Com4jObject {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(11)
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(14)
  void setNumberFormat(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLinked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getNumberFormatLinked();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLinked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(16)
  void setNumberFormatLinked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(17)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(18)
  void setNumberFormatLocal(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTickLabelOrientation
   */

  @VTID(19)
  com.exceljava.com4j.excel.XlTickLabelOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTickLabelOrientation parameter.
   */

  @VTID(20)
  void setOrientation(
    com.exceljava.com4j.excel.XlTickLabelOrientation rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(22)
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(23)
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoScaleFont"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(24)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAutoScaleFont();


  /**
   * <p>
   * Setter method for the COM property "AutoScaleFont"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(25)
  void setAutoScaleFont(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Depth"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(26)
  int getDepth();


  /**
   * <p>
   * Getter method for the COM property "Offset"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(27)
  int getOffset();


  /**
   * <p>
   * Setter method for the COM property "Offset"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(28)
  void setOffset(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Alignment"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(29)
  int getAlignment();


  /**
   * <p>
   * Setter method for the COM property "Alignment"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(30)
  void setAlignment(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MultiLevel"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getMultiLevel();


  /**
   * <p>
   * Setter method for the COM property "MultiLevel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setMultiLevel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartFormat
   */

  @VTID(33)
  com.exceljava.com4j.excel.ChartFormat getFormat();


  // Properties:
}
