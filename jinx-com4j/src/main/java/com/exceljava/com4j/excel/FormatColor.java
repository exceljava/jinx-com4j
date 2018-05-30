package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface FormatColor extends Com4jObject {
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
   * Getter method for the COM property "Color"
   * </p>
   */

  @DISPID(99)
  @PropGet
  java.lang.Object getColor();


  /**
   * <p>
   * Setter method for the COM property "Color"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(99)
  @PropPut
  void setColor(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ColorIndex"
   * </p>
   */

  @DISPID(97)
  @PropGet
  com.exceljava.com4j.excel.XlColorIndex getColorIndex();


  /**
   * <p>
   * Setter method for the COM property "ColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @DISPID(97)
  @PropPut
  void setColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "ThemeColor"
   * </p>
   */

  @DISPID(2365)
  @PropGet
  java.lang.Object getThemeColor();


  /**
   * <p>
   * Setter method for the COM property "ThemeColor"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2365)
  @PropPut
  void setThemeColor(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TintAndShade"
   * </p>
   */

  @DISPID(2366)
  @PropGet
  java.lang.Object getTintAndShade();


  /**
   * <p>
   * Setter method for the COM property "TintAndShade"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2366)
  @PropPut
  void setTintAndShade(
    java.lang.Object rhs);


  // Properties:
}
