package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C170E-0000-0000-C000-000000000046}")
public interface IMsoTrendline extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Backward"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(8)
  double getBackward();


  /**
   * <p>
   * Setter method for the COM property "Backward"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(9)
  void setBackward(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoBorder
   */

  @VTID(10)
  com.exceljava.com4j.office.IMsoBorder getBorder();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearFormats();


  /**
   * <p>
   * Getter method for the COM property "DataLabel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDataLabel
   */

  @VTID(12)
  com.exceljava.com4j.office.IMsoDataLabel getDataLabel();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "DisplayEquation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getDisplayEquation();


  /**
   * <p>
   * Setter method for the COM property "DisplayEquation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setDisplayEquation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayRSquared"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getDisplayRSquared();


  /**
   * <p>
   * Setter method for the COM property "DisplayRSquared"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setDisplayRSquared(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(18)
  double getForward();


  /**
   * <p>
   * Setter method for the COM property "Forward"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(19)
  void setForward(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(20)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Intercept"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(21)
  double getIntercept();


  /**
   * <p>
   * Setter method for the COM property "Intercept"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(22)
  void setIntercept(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InterceptIsAuto"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(23)
  boolean getInterceptIsAuto();


  /**
   * <p>
   * Setter method for the COM property "InterceptIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(24)
  void setInterceptIsAuto(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(25)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(26)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NameIsAuto"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(27)
  boolean getNameIsAuto();


  /**
   * <p>
   * Setter method for the COM property "NameIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(28)
  void setNameIsAuto(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(29)
  int getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(30)
  void setOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Period"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(31)
  int getPeriod();


  /**
   * <p>
   * Setter method for the COM property "Period"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(32)
  void setPeriod(
    int rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlTrendlineType
   */

  @VTID(34)
  com.exceljava.com4j.office.XlTrendlineType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.XlTrendlineType parameter.
   */

  @VTID(35)
  void setType(
    com.exceljava.com4j.office.XlTrendlineType rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartFormat
   */

  @VTID(36)
  com.exceljava.com4j.office.IMsoChartFormat getFormat();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(37)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(38)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "Backward2"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(39)
  double getBackward2();


  /**
   * <p>
   * Setter method for the COM property "Backward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(40)
  void setBackward2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward2"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(41)
  double getForward2();


  /**
   * <p>
   * Setter method for the COM property "Forward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(42)
  void setForward2(
    double rhs);


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @VTID(43)
  void setProperty(
    java.lang.String bstrId,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String bstrId);


  // Properties:
}
