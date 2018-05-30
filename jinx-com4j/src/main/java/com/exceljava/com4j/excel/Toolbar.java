package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Toolbar extends Com4jObject {
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
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  int getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  int getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(127)
  @PropPut
  void setLeft(
    int rhs);


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
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  int getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(133)
  @PropPut
  void setPosition(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Protection"
   * </p>
   */

  @DISPID(176)
  @PropGet
  com.exceljava.com4j.excel.XlToolbarProtection getProtection();


  /**
   * <p>
   * Setter method for the COM property "Protection"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlToolbarProtection parameter.
   */

  @DISPID(176)
  @PropPut
  void setProtection(
    com.exceljava.com4j.excel.XlToolbarProtection rhs);


  /**
   */

  @DISPID(555)
  void reset();


  /**
   * <p>
   * Getter method for the COM property "ToolbarButtons"
   * </p>
   */

  @DISPID(964)
  @PropGet
  com.exceljava.com4j.excel.ToolbarButtons getToolbarButtons();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  int getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(126)
  @PropPut
  void setTop(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    boolean rhs);


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
