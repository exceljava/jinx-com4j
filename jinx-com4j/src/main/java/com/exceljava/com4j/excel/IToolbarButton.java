package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002085E-0001-0000-C000-000000000046}")
public interface IToolbarButton extends Com4jObject {
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
   * Getter method for the COM property "BuiltIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "BuiltInFace"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getBuiltInFace();


  /**
   * <p>
   * Setter method for the COM property "BuiltInFace"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(12)
  void setBuiltInFace(
    boolean rhs);


  /**
   * @param toolbar Mandatory com.exceljava.com4j.excel.Toolbar parameter.
   * @param before Mandatory int parameter.
   */

  @VTID(13)
  void copy(
    com.exceljava.com4j.excel.Toolbar toolbar,
    int before);


  /**
   */

  @VTID(14)
  void copyFace();


  /**
   */

  @VTID(15)
  void delete();


  /**
   */

  @VTID(16)
  void edit();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HelpContextID"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(19)
  int getHelpContextID();


  /**
   * <p>
   * Setter method for the COM property "HelpContextID"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(20)
  void setHelpContextID(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "HelpFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(21)
  java.lang.String getHelpFile();


  /**
   * <p>
   * Setter method for the COM property "HelpFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(22)
  void setHelpFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(23)
  int getID();


  /**
   * <p>
   * Getter method for the COM property "IsGap"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getIsGap();


  /**
   * @param toolbar Mandatory com.exceljava.com4j.excel.Toolbar parameter.
   * @param before Mandatory int parameter.
   */

  @VTID(25)
  void move(
    com.exceljava.com4j.excel.Toolbar toolbar,
    int before);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(26)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(27)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(28)
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(29)
  void setOnAction(
    java.lang.String rhs);


  /**
   */

  @VTID(30)
  void pasteFace();


  /**
   * <p>
   * Getter method for the COM property "Pushed"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getPushed();


  /**
   * <p>
   * Setter method for the COM property "Pushed"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setPushed(
    boolean rhs);


  /**
   */

  @VTID(33)
  void reset();


  /**
   * <p>
   * Getter method for the COM property "StatusBar"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(34)
  java.lang.String getStatusBar();


  /**
   * <p>
   * Setter method for the COM property "StatusBar"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(35)
  void setStatusBar(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(36)
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(37)
  void setWidth(
    int rhs);


  // Properties:
}
