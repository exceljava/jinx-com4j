package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208BE-0001-0000-C000-000000000046}")
public interface ITrendline extends Com4jObject {
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
   * Getter method for the COM property "Backward"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getBackward();


  /**
   * <p>
   * Setter method for the COM property "Backward"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(11)
  void setBackward(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Border
   */

  @VTID(12)
  com.exceljava.com4j.excel.Border getBorder();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearFormats();


  /**
   * <p>
   * Getter method for the COM property "DataLabel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DataLabel
   */

  @VTID(14)
  com.exceljava.com4j.excel.DataLabel getDataLabel();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "DisplayEquation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getDisplayEquation();


  /**
   * <p>
   * Setter method for the COM property "DisplayEquation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setDisplayEquation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayRSquared"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getDisplayRSquared();


  /**
   * <p>
   * Setter method for the COM property "DisplayRSquared"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setDisplayRSquared(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(20)
  int getForward();


  /**
   * <p>
   * Setter method for the COM property "Forward"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(21)
  void setForward(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(22)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Intercept"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(23)
  double getIntercept();


  /**
   * <p>
   * Setter method for the COM property "Intercept"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(24)
  void setIntercept(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InterceptIsAuto"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(25)
  boolean getInterceptIsAuto();


  /**
   * <p>
   * Setter method for the COM property "InterceptIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(26)
  void setInterceptIsAuto(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(28)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NameIsAuto"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getNameIsAuto();


  /**
   * <p>
   * Setter method for the COM property "NameIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(30)
  void setNameIsAuto(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(31)
  int getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(32)
  void setOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Period"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(33)
  int getPeriod();


  /**
   * <p>
   * Setter method for the COM property "Period"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(34)
  void setPeriod(
    int rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(35)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTrendlineType
   */

  @VTID(36)
  com.exceljava.com4j.excel.XlTrendlineType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTrendlineType parameter.
   */

  @VTID(37)
  void setType(
    com.exceljava.com4j.excel.XlTrendlineType rhs);


  /**
   * <p>
   * Getter method for the COM property "Backward2"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(38)
  double getBackward2();


  /**
   * <p>
   * Setter method for the COM property "Backward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(39)
  void setBackward2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward2"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(40)
  double getForward2();


  /**
   * <p>
   * Setter method for the COM property "Forward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(41)
  void setForward2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartFormat
   */

  @VTID(42)
  com.exceljava.com4j.excel.ChartFormat getFormat();


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @VTID(43)
  void setProperty(
    java.lang.String id,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String id);


  // Properties:
}
