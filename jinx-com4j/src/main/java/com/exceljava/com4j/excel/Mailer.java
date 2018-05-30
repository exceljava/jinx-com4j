package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Mailer extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "BCCRecipients"
   * </p>
   */

  @DISPID(983)
  @PropGet
  java.lang.Object getBCCRecipients();


  /**
   * <p>
   * Setter method for the COM property "BCCRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(983)
  @PropPut
  void setBCCRecipients(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CCRecipients"
   * </p>
   */

  @DISPID(982)
  @PropGet
  java.lang.Object getCCRecipients();


  /**
   * <p>
   * Setter method for the COM property "CCRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(982)
  @PropPut
  void setCCRecipients(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Enclosures"
   * </p>
   */

  @DISPID(984)
  @PropGet
  java.lang.Object getEnclosures();


  /**
   * <p>
   * Setter method for the COM property "Enclosures"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(984)
  @PropPut
  void setEnclosures(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Received"
   * </p>
   */

  @DISPID(986)
  @PropGet
  boolean getReceived();


  /**
   * <p>
   * Getter method for the COM property "SendDateTime"
   * </p>
   */

  @DISPID(987)
  @PropGet
  java.util.Date getSendDateTime();


  /**
   * <p>
   * Getter method for the COM property "Sender"
   * </p>
   */

  @DISPID(988)
  @PropGet
  java.lang.String getSender();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   */

  @DISPID(953)
  @PropGet
  java.lang.String getSubject();


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(953)
  @PropPut
  void setSubject(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ToRecipients"
   * </p>
   */

  @DISPID(981)
  @PropGet
  java.lang.Object getToRecipients();


  /**
   * <p>
   * Setter method for the COM property "ToRecipients"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(981)
  @PropPut
  void setToRecipients(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "WhichAddress"
   * </p>
   */

  @DISPID(974)
  @PropGet
  java.lang.Object getWhichAddress();


  /**
   * <p>
   * Setter method for the COM property "WhichAddress"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(974)
  @PropPut
  void setWhichAddress(
    java.lang.Object rhs);


  // Properties:
}
