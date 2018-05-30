package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0398-0000-0000-C000-000000000046}")
public interface TextFrame2 extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "MarginBottom"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(10)
  float getMarginBottom();


  /**
   * <p>
   * Setter method for the COM property "MarginBottom"
   * </p>
   * @param marginBottom Mandatory float parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(11)
  void setMarginBottom(
    float marginBottom);


  /**
   * <p>
   * Getter method for the COM property "MarginLeft"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(12)
  float getMarginLeft();


  /**
   * <p>
   * Setter method for the COM property "MarginLeft"
   * </p>
   * @param marginLeft Mandatory float parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(13)
  void setMarginLeft(
    float marginLeft);


  /**
   * <p>
   * Getter method for the COM property "MarginRight"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(14)
  float getMarginRight();


  /**
   * <p>
   * Setter method for the COM property "MarginRight"
   * </p>
   * @param marginRight Mandatory float parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(15)
  void setMarginRight(
    float marginRight);


  /**
   * <p>
   * Getter method for the COM property "MarginTop"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(16)
  float getMarginTop();


  /**
   * <p>
   * Setter method for the COM property "MarginTop"
   * </p>
   * @param marginTop Mandatory float parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(17)
  void setMarginTop(
    float marginTop);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTextOrientation
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoTextOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param orientation Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(19)
  void setOrientation(
    com.exceljava.com4j.office.MsoTextOrientation orientation);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAnchor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoHorizontalAnchor
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.MsoHorizontalAnchor getHorizontalAnchor();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAnchor"
   * </p>
   * @param horizontalAnchor Mandatory com.exceljava.com4j.office.MsoHorizontalAnchor parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(21)
  void setHorizontalAnchor(
    com.exceljava.com4j.office.MsoHorizontalAnchor horizontalAnchor);


  /**
   * <p>
   * Getter method for the COM property "VerticalAnchor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoVerticalAnchor
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.MsoVerticalAnchor getVerticalAnchor();


  /**
   * <p>
   * Setter method for the COM property "VerticalAnchor"
   * </p>
   * @param verticalAnchor Mandatory com.exceljava.com4j.office.MsoVerticalAnchor parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(23)
  void setVerticalAnchor(
    com.exceljava.com4j.office.MsoVerticalAnchor verticalAnchor);


  /**
   * <p>
   * Getter method for the COM property "PathFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPathFormat
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoPathFormat getPathFormat();


  /**
   * <p>
   * Setter method for the COM property "PathFormat"
   * </p>
   * @param pathFormat Mandatory com.exceljava.com4j.office.MsoPathFormat parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(25)
  void setPathFormat(
    com.exceljava.com4j.office.MsoPathFormat pathFormat);


  /**
   * <p>
   * Getter method for the COM property "WarpFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoWarpFormat
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.office.MsoWarpFormat getWarpFormat();


  /**
   * <p>
   * Setter method for the COM property "WarpFormat"
   * </p>
   * @param warpFormat Mandatory com.exceljava.com4j.office.MsoWarpFormat parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(27)
  void setWarpFormat(
    com.exceljava.com4j.office.MsoWarpFormat warpFormat);


  /**
   * <p>
   * Getter method for the COM property "WordArtformat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetTextEffect
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.MsoPresetTextEffect getWordArtformat();


  /**
   * <p>
   * Setter method for the COM property "WordArtformat"
   * </p>
   * @param wordArtformat Mandatory com.exceljava.com4j.office.MsoPresetTextEffect parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(29)
  void setWordArtformat(
    com.exceljava.com4j.office.MsoPresetTextEffect wordArtformat);


  /**
   * <p>
   * Getter method for the COM property "WordWrap"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.MsoTriState getWordWrap();


  /**
   * <p>
   * Setter method for the COM property "WordWrap"
   * </p>
   * @param wordWrap Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(31)
  void setWordWrap(
    com.exceljava.com4j.office.MsoTriState wordWrap);


  /**
   * <p>
   * Getter method for the COM property "AutoSize"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoSize
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(32)
  com.exceljava.com4j.office.MsoAutoSize getAutoSize();


  /**
   * <p>
   * Setter method for the COM property "AutoSize"
   * </p>
   * @param autoSize Mandatory com.exceljava.com4j.office.MsoAutoSize parameter.
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(33)
  void setAutoSize(
    com.exceljava.com4j.office.MsoAutoSize autoSize);


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThreeDFormat
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(34)
  com.exceljava.com4j.office.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "HasText"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(35)
  com.exceljava.com4j.office.MsoTriState getHasText();


  /**
   * <p>
   * Getter method for the COM property "TextRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(36)
  com.exceljava.com4j.office.TextRange2 getTextRange();


  /**
   * <p>
   * Getter method for the COM property "Column"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextColumn2
   */

  @DISPID(115) //= 0x73. The runtime will prefer the VTID if present
  @VTID(37)
  com.exceljava.com4j.office.TextColumn2 getColumn();


  /**
   * <p>
   * Getter method for the COM property "Ruler"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Ruler2
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.Ruler2 getRuler();


  /**
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(39)
  void deleteText();


  /**
   * <p>
   * Getter method for the COM property "NoTextRotation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(40)
  com.exceljava.com4j.office.MsoTriState getNoTextRotation();


  /**
   * <p>
   * Setter method for the COM property "NoTextRotation"
   * </p>
   * @param noTextRotation Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(41)
  void setNoTextRotation(
    com.exceljava.com4j.office.MsoTriState noTextRotation);


  // Properties:
}
