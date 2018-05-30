package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ChartFormat extends Com4jObject {
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
   * Getter method for the COM property "Fill"
   * </p>
   */

  @DISPID(1663)
  @PropGet
  com.exceljava.com4j.excel.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   */

  @DISPID(2663)
  @PropGet
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   */

  @DISPID(817)
  @PropGet
  com.exceljava.com4j.excel.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   */

  @DISPID(1631)
  @PropGet
  com.exceljava.com4j.excel.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   */

  @DISPID(103)
  @PropGet
  com.exceljava.com4j.excel.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   */

  @DISPID(2662)
  @PropGet
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   */

  @DISPID(2659)
  @PropGet
  com.exceljava.com4j.excel.TextFrame2 getTextFrame2();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   */

  @DISPID(1703)
  @PropGet
  com.exceljava.com4j.excel.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   */

  @DISPID(1691)
  @PropGet
  com.exceljava.com4j.excel.Adjustments getAdjustments();


  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   */

  @DISPID(1693)
  @PropGet
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @DISPID(1693)
  @PropPut
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType rhs);


  // Properties:
}
