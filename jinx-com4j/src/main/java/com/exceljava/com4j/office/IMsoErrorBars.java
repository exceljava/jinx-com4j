package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1721-0000-0000-C000-000000000046}")
public interface IMsoErrorBars extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(8)
  java.lang.String getName();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(9)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


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
  java.lang.Object delete();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearFormats();


  /**
   * <p>
   * Getter method for the COM property "EndStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlEndStyleCap
   */

  @VTID(13)
  com.exceljava.com4j.office.XlEndStyleCap getEndStyle();


  /**
   * <p>
   * Setter method for the COM property "EndStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.XlEndStyleCap parameter.
   */

  @VTID(14)
  void setEndStyle(
    com.exceljava.com4j.office.XlEndStyleCap rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartFormat
   */

  @VTID(15)
  com.exceljava.com4j.office.IMsoChartFormat getFormat();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(17)
  int getCreator();


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @VTID(18)
  void setProperty(
    java.lang.String bstrId,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(19)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String bstrId);


  // Properties:
}
