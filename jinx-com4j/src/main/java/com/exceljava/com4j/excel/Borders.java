package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Borders extends Com4jObject,Iterable<Com4jObject> {
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
  java.lang.Object getColorIndex();


  /**
   * <p>
   * Setter method for the COM property "ColorIndex"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(97)
  @PropPut
  void setColorIndex(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory com.exceljava.com4j.excel.XlBordersIndex parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.Border getItem(
    com.exceljava.com4j.excel.XlBordersIndex index);


  /**
   * <p>
   * Getter method for the COM property "LineStyle"
   * </p>
   */

  @DISPID(119)
  @PropGet
  java.lang.Object getLineStyle();


  /**
   * <p>
   * Setter method for the COM property "LineStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(119)
  @PropPut
  void setLineStyle(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.Object getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Weight"
   * </p>
   */

  @DISPID(120)
  @PropGet
  java.lang.Object getWeight();


  /**
   * <p>
   * Setter method for the COM property "Weight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(120)
  @PropPut
  void setWeight(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory com.exceljava.com4j.excel.XlBordersIndex parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.Border get_Default(
    com.exceljava.com4j.excel.XlBordersIndex index);


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
