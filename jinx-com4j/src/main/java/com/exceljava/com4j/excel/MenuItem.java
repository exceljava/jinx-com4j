package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface MenuItem extends Com4jObject {
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
   * Getter method for the COM property "Caption"
   * </p>
   */

  @DISPID(139)
  @PropGet
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(139)
  @PropPut
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Checked"
   * </p>
   */

  @DISPID(599)
  @PropGet
  boolean getChecked();


  /**
   * <p>
   * Setter method for the COM property "Checked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(599)
  @PropPut
  void setChecked(
    boolean rhs);


  /**
   */

  @DISPID(117)
  void delete();


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
   * Getter method for the COM property "Index"
   * </p>
   */

  @DISPID(486)
  @PropGet
  int getIndex();


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


  // Properties:
}
