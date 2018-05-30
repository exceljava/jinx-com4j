package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024459-0001-0000-C000-000000000046}")
public interface IGraphic extends Com4jObject {
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
   * Getter method for the COM property "Brightness"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(10)
  float getBrightness();


  /**
   * <p>
   * Setter method for the COM property "Brightness"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(11)
  void setBrightness(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPictureColorType
   */

  @VTID(12)
  com.exceljava.com4j.office.MsoPictureColorType getColorType();


  /**
   * <p>
   * Setter method for the COM property "ColorType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoPictureColorType parameter.
   */

  @VTID(13)
  void setColorType(
    com.exceljava.com4j.office.MsoPictureColorType rhs);


  /**
   * <p>
   * Getter method for the COM property "Contrast"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(14)
  float getContrast();


  /**
   * <p>
   * Setter method for the COM property "Contrast"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(15)
  void setContrast(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropBottom"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(16)
  float getCropBottom();


  /**
   * <p>
   * Setter method for the COM property "CropBottom"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(17)
  void setCropBottom(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropLeft"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(18)
  float getCropLeft();


  /**
   * <p>
   * Setter method for the COM property "CropLeft"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(19)
  void setCropLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropRight"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(20)
  float getCropRight();


  /**
   * <p>
   * Setter method for the COM property "CropRight"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(21)
  void setCropRight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropTop"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(22)
  float getCropTop();


  /**
   * <p>
   * Setter method for the COM property "CropTop"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(23)
  void setCropTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Filename"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(24)
  java.lang.String getFilename();


  /**
   * <p>
   * Setter method for the COM property "Filename"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(25)
  void setFilename(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(26)
  float getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(27)
  void setHeight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "LockAspectRatio"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(28)
  com.exceljava.com4j.office.MsoTriState getLockAspectRatio();


  /**
   * <p>
   * Setter method for the COM property "LockAspectRatio"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(29)
  void setLockAspectRatio(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(30)
  float getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(31)
  void setWidth(
    float rhs);


  // Properties:
}
