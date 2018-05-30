package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0317-0000-0000-C000-000000000046}")
public interface LineFormat extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "BackColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ColorFormat
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.ColorFormat getBackColor();


  /**
   * <p>
   * Setter method for the COM property "BackColor"
   * </p>
   * @param backColor Mandatory com.exceljava.com4j.office.ColorFormat parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(11)
  void setBackColor(
    com.exceljava.com4j.office.ColorFormat backColor);


  /**
   * <p>
   * Getter method for the COM property "BeginArrowheadLength"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadLength
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.MsoArrowheadLength getBeginArrowheadLength();


  /**
   * <p>
   * Setter method for the COM property "BeginArrowheadLength"
   * </p>
   * @param beginArrowheadLength Mandatory com.exceljava.com4j.office.MsoArrowheadLength parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(13)
  void setBeginArrowheadLength(
    com.exceljava.com4j.office.MsoArrowheadLength beginArrowheadLength);


  /**
   * <p>
   * Getter method for the COM property "BeginArrowheadStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadStyle
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoArrowheadStyle getBeginArrowheadStyle();


  /**
   * <p>
   * Setter method for the COM property "BeginArrowheadStyle"
   * </p>
   * @param beginArrowheadStyle Mandatory com.exceljava.com4j.office.MsoArrowheadStyle parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(15)
  void setBeginArrowheadStyle(
    com.exceljava.com4j.office.MsoArrowheadStyle beginArrowheadStyle);


  /**
   * <p>
   * Getter method for the COM property "BeginArrowheadWidth"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadWidth
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.MsoArrowheadWidth getBeginArrowheadWidth();


  /**
   * <p>
   * Setter method for the COM property "BeginArrowheadWidth"
   * </p>
   * @param beginArrowheadWidth Mandatory com.exceljava.com4j.office.MsoArrowheadWidth parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(17)
  void setBeginArrowheadWidth(
    com.exceljava.com4j.office.MsoArrowheadWidth beginArrowheadWidth);


  /**
   * <p>
   * Getter method for the COM property "DashStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoLineDashStyle
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoLineDashStyle getDashStyle();


  /**
   * <p>
   * Setter method for the COM property "DashStyle"
   * </p>
   * @param dashStyle Mandatory com.exceljava.com4j.office.MsoLineDashStyle parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(19)
  void setDashStyle(
    com.exceljava.com4j.office.MsoLineDashStyle dashStyle);


  /**
   * <p>
   * Getter method for the COM property "EndArrowheadLength"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadLength
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.MsoArrowheadLength getEndArrowheadLength();


  /**
   * <p>
   * Setter method for the COM property "EndArrowheadLength"
   * </p>
   * @param endArrowheadLength Mandatory com.exceljava.com4j.office.MsoArrowheadLength parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(21)
  void setEndArrowheadLength(
    com.exceljava.com4j.office.MsoArrowheadLength endArrowheadLength);


  /**
   * <p>
   * Getter method for the COM property "EndArrowheadStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadStyle
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.MsoArrowheadStyle getEndArrowheadStyle();


  /**
   * <p>
   * Setter method for the COM property "EndArrowheadStyle"
   * </p>
   * @param endArrowheadStyle Mandatory com.exceljava.com4j.office.MsoArrowheadStyle parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(23)
  void setEndArrowheadStyle(
    com.exceljava.com4j.office.MsoArrowheadStyle endArrowheadStyle);


  /**
   * <p>
   * Getter method for the COM property "EndArrowheadWidth"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoArrowheadWidth
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoArrowheadWidth getEndArrowheadWidth();


  /**
   * <p>
   * Setter method for the COM property "EndArrowheadWidth"
   * </p>
   * @param endArrowheadWidth Mandatory com.exceljava.com4j.office.MsoArrowheadWidth parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(25)
  void setEndArrowheadWidth(
    com.exceljava.com4j.office.MsoArrowheadWidth endArrowheadWidth);


  /**
   * <p>
   * Getter method for the COM property "ForeColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ColorFormat
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.office.ColorFormat getForeColor();


  /**
   * <p>
   * Setter method for the COM property "ForeColor"
   * </p>
   * @param foreColor Mandatory com.exceljava.com4j.office.ColorFormat parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(27)
  void setForeColor(
    com.exceljava.com4j.office.ColorFormat foreColor);


  /**
   * <p>
   * Getter method for the COM property "Pattern"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPatternType
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.MsoPatternType getPattern();


  /**
   * <p>
   * Setter method for the COM property "Pattern"
   * </p>
   * @param pattern Mandatory com.exceljava.com4j.office.MsoPatternType parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(29)
  void setPattern(
    com.exceljava.com4j.office.MsoPatternType pattern);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoLineStyle
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.MsoLineStyle getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param style Mandatory com.exceljava.com4j.office.MsoLineStyle parameter.
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(31)
  void setStyle(
    com.exceljava.com4j.office.MsoLineStyle style);


  /**
   * <p>
   * Getter method for the COM property "Transparency"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(32)
  float getTransparency();


  /**
   * <p>
   * Setter method for the COM property "Transparency"
   * </p>
   * @param transparency Mandatory float parameter.
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(33)
  void setTransparency(
    float transparency);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(34)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param visible Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(35)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState visible);


  /**
   * <p>
   * Getter method for the COM property "Weight"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(36)
  float getWeight();


  /**
   * <p>
   * Setter method for the COM property "Weight"
   * </p>
   * @param weight Mandatory float parameter.
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(37)
  void setWeight(
    float weight);


  /**
   * <p>
   * Getter method for the COM property "InsetPen"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.MsoTriState getInsetPen();


  /**
   * <p>
   * Setter method for the COM property "InsetPen"
   * </p>
   * @param insetPen Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(39)
  void setInsetPen(
    com.exceljava.com4j.office.MsoTriState insetPen);


  // Properties:
}
