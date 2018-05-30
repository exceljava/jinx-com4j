package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208D1-0001-0000-C000-000000000046}")
public interface IMailer extends Com4jObject {
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
   * Getter method for the COM property "BCCRecipients"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getBCCRecipients();


  /**
   * <p>
   * Setter method for the COM property "BCCRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(11)
  void setBCCRecipients(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CCRecipients"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCCRecipients();


  /**
   * <p>
   * Setter method for the COM property "CCRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(13)
  void setCCRecipients(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Enclosures"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getEnclosures();


  /**
   * <p>
   * Setter method for the COM property "Enclosures"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setEnclosures(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Received"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getReceived();


  /**
   * <p>
   * Getter method for the COM property "SendDateTime"
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @VTID(17)
  java.util.Date getSendDateTime();


  /**
   * <p>
   * Getter method for the COM property "Sender"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(18)
  java.lang.String getSender();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(19)
  java.lang.String getSubject();


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(20)
  void setSubject(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ToRecipients"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getToRecipients();


  /**
   * <p>
   * Setter method for the COM property "ToRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(22)
  void setToRecipients(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "WhichAddress"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getWhichAddress();


  /**
   * <p>
   * Setter method for the COM property "WhichAddress"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(24)
  void setWhichAddress(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  // Properties:
}
