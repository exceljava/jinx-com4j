package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002085C-0001-0000-C000-000000000046}")
public interface IToolbar extends Com4jObject {
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
   */

  @VTID(11)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(13)
  void setHeight(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
  int getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(15)
  void setLeft(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(16)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(17)
  int getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(18)
  void setPosition(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Protection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlToolbarProtection
   */

  @VTID(19)
  com.exceljava.com4j.excel.XlToolbarProtection getProtection();


  /**
   * <p>
   * Setter method for the COM property "Protection"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlToolbarProtection parameter.
   */

  @VTID(20)
  void setProtection(
    com.exceljava.com4j.excel.XlToolbarProtection rhs);


  /**
   */

  @VTID(21)
  void reset();


  /**
   * <p>
   * Getter method for the COM property "ToolbarButtons"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButtons
   */

  @VTID(22)
  com.exceljava.com4j.excel.ToolbarButtons getToolbarButtons();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(23)
  int getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(24)
  void setTop(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(25)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(26)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(27)
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(28)
  void setWidth(
    int rhs);


  // Properties:
}
