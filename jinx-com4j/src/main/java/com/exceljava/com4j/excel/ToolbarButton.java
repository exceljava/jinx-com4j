package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ToolbarButton extends Com4jObject {
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
   * Getter method for the COM property "BuiltIn"
   * </p>
   */

  @DISPID(553)
  @PropGet
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "BuiltInFace"
   * </p>
   */

  @DISPID(554)
  @PropGet
  boolean getBuiltInFace();


  /**
   * <p>
   * Setter method for the COM property "BuiltInFace"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(554)
  @PropPut
  void setBuiltInFace(
    boolean rhs);


  /**
   * @param toolbar Mandatory com.exceljava.com4j.excel.Toolbar parameter.
   * @param before Mandatory int parameter.
   */

  @DISPID(551)
  void copy(
    com.exceljava.com4j.excel.Toolbar toolbar,
    int before);


  /**
   */

  @DISPID(966)
  void copyFace();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   */

  @DISPID(562)
  void edit();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   */

  @DISPID(600)
  @PropGet
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(600)
  @PropPut
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HelpContextID"
   * </p>
   */

  @DISPID(355)
  @PropGet
  int getHelpContextID();


  /**
   * <p>
   * Setter method for the COM property "HelpContextID"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(355)
  @PropPut
  void setHelpContextID(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "HelpFile"
   * </p>
   */

  @DISPID(360)
  @PropGet
  java.lang.String getHelpFile();


  /**
   * <p>
   * Setter method for the COM property "HelpFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(360)
  @PropPut
  void setHelpFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   */

  @DISPID(570)
  @PropGet
  int getID();


  /**
   * <p>
   * Getter method for the COM property "IsGap"
   * </p>
   */

  @DISPID(561)
  @PropGet
  boolean getIsGap();


  /**
   * @param toolbar Mandatory com.exceljava.com4j.excel.Toolbar parameter.
   * @param before Mandatory int parameter.
   */

  @DISPID(637)
  void move(
    com.exceljava.com4j.excel.Toolbar toolbar,
    int before);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   */

  @DISPID(596)
  @PropGet
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(596)
  @PropPut
  void setOnAction(
    java.lang.String rhs);


  /**
   */

  @DISPID(967)
  void pasteFace();


  /**
   * <p>
   * Getter method for the COM property "Pushed"
   * </p>
   */

  @DISPID(560)
  @PropGet
  boolean getPushed();


  /**
   * <p>
   * Setter method for the COM property "Pushed"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(560)
  @PropPut
  void setPushed(
    boolean rhs);


  /**
   */

  @DISPID(555)
  void reset();


  /**
   * <p>
   * Getter method for the COM property "StatusBar"
   * </p>
   */

  @DISPID(386)
  @PropGet
  java.lang.String getStatusBar();


  /**
   * <p>
   * Setter method for the COM property "StatusBar"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(386)
  @PropPut
  void setStatusBar(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    int rhs);


  // Properties:
}
