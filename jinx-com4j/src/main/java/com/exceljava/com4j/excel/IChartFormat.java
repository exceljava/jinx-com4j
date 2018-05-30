package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244B2-0001-0000-C000-000000000046}")
public interface IChartFormat extends Com4jObject {
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
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.FillFormat
   */

  @VTID(10)
  com.exceljava.com4j.excel.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GlowFormat
   */

  @VTID(11)
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.LineFormat
   */

  @VTID(12)
  com.exceljava.com4j.excel.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PictureFormat
   */

  @VTID(13)
  com.exceljava.com4j.excel.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShadowFormat
   */

  @VTID(14)
  com.exceljava.com4j.excel.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SoftEdgeFormat
   */

  @VTID(15)
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextFrame2
   */

  @VTID(16)
  com.exceljava.com4j.excel.TextFrame2 getTextFrame2();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ThreeDFormat
   */

  @VTID(17)
  com.exceljava.com4j.excel.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Adjustments
   */

  @VTID(18)
  com.exceljava.com4j.excel.Adjustments getAdjustments();


  @VTID(18)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.Adjustments.class})
  float getAdjustments(
    int index);

  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoShapeType
   */

  @VTID(19)
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @VTID(20)
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType rhs);


  // Properties:
}
