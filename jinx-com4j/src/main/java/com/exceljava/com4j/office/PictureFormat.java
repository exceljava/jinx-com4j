package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C031A-0000-0000-C000-000000000046}")
public interface PictureFormat extends com.exceljava.com4j.office._IMsoDispObj {
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
   * @param increment Mandatory float parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  void incrementBrightness(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(11)
  void incrementContrast(
    float increment);


  /**
   * <p>
   * Getter method for the COM property "Brightness"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(12)
  float getBrightness();


  /**
   * <p>
   * Setter method for the COM property "Brightness"
   * </p>
   * @param brightness Mandatory float parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(13)
  void setBrightness(
    float brightness);


  /**
   * <p>
   * Getter method for the COM property "ColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPictureColorType
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoPictureColorType getColorType();


  /**
   * <p>
   * Setter method for the COM property "ColorType"
   * </p>
   * @param colorType Mandatory com.exceljava.com4j.office.MsoPictureColorType parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(15)
  void setColorType(
    com.exceljava.com4j.office.MsoPictureColorType colorType);


  /**
   * <p>
   * Getter method for the COM property "Contrast"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(16)
  float getContrast();


  /**
   * <p>
   * Setter method for the COM property "Contrast"
   * </p>
   * @param contrast Mandatory float parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(17)
  void setContrast(
    float contrast);


  /**
   * <p>
   * Getter method for the COM property "CropBottom"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(18)
  float getCropBottom();


  /**
   * <p>
   * Setter method for the COM property "CropBottom"
   * </p>
   * @param cropBottom Mandatory float parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(19)
  void setCropBottom(
    float cropBottom);


  /**
   * <p>
   * Getter method for the COM property "CropLeft"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(20)
  float getCropLeft();


  /**
   * <p>
   * Setter method for the COM property "CropLeft"
   * </p>
   * @param cropLeft Mandatory float parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(21)
  void setCropLeft(
    float cropLeft);


  /**
   * <p>
   * Getter method for the COM property "CropRight"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(22)
  float getCropRight();


  /**
   * <p>
   * Setter method for the COM property "CropRight"
   * </p>
   * @param cropRight Mandatory float parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(23)
  void setCropRight(
    float cropRight);


  /**
   * <p>
   * Getter method for the COM property "CropTop"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(24)
  float getCropTop();


  /**
   * <p>
   * Setter method for the COM property "CropTop"
   * </p>
   * @param cropTop Mandatory float parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(25)
  void setCropTop(
    float cropTop);


  /**
   * <p>
   * Getter method for the COM property "TransparencyColor"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(26)
  int getTransparencyColor();


  /**
   * <p>
   * Setter method for the COM property "TransparencyColor"
   * </p>
   * @param transparencyColor Mandatory int parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(27)
  void setTransparencyColor(
    int transparencyColor);


  /**
   * <p>
   * Getter method for the COM property "TransparentBackground"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.MsoTriState getTransparentBackground();


  /**
   * <p>
   * Setter method for the COM property "TransparentBackground"
   * </p>
   * @param transparentBackground Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(29)
  void setTransparentBackground(
    com.exceljava.com4j.office.MsoTriState transparentBackground);


  /**
   * <p>
   * Getter method for the COM property "Crop"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Crop
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.Crop getCrop();


  // Properties:
}
