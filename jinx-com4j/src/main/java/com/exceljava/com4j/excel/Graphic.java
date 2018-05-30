package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Graphic extends Com4jObject {
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
   * Getter method for the COM property "Brightness"
   * </p>
   */

  @DISPID(2194)
  @PropGet
  float getBrightness();


  /**
   * <p>
   * Setter method for the COM property "Brightness"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2194)
  @PropPut
  void setBrightness(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ColorType"
   * </p>
   */

  @DISPID(2195)
  @PropGet
  com.exceljava.com4j.office.MsoPictureColorType getColorType();


  /**
   * <p>
   * Setter method for the COM property "ColorType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoPictureColorType parameter.
   */

  @DISPID(2195)
  @PropPut
  void setColorType(
    com.exceljava.com4j.office.MsoPictureColorType rhs);


  /**
   * <p>
   * Getter method for the COM property "Contrast"
   * </p>
   */

  @DISPID(2196)
  @PropGet
  float getContrast();


  /**
   * <p>
   * Setter method for the COM property "Contrast"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2196)
  @PropPut
  void setContrast(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropBottom"
   * </p>
   */

  @DISPID(2197)
  @PropGet
  float getCropBottom();


  /**
   * <p>
   * Setter method for the COM property "CropBottom"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2197)
  @PropPut
  void setCropBottom(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropLeft"
   * </p>
   */

  @DISPID(2198)
  @PropGet
  float getCropLeft();


  /**
   * <p>
   * Setter method for the COM property "CropLeft"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2198)
  @PropPut
  void setCropLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropRight"
   * </p>
   */

  @DISPID(2199)
  @PropGet
  float getCropRight();


  /**
   * <p>
   * Setter method for the COM property "CropRight"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2199)
  @PropPut
  void setCropRight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "CropTop"
   * </p>
   */

  @DISPID(2200)
  @PropGet
  float getCropTop();


  /**
   * <p>
   * Setter method for the COM property "CropTop"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2200)
  @PropPut
  void setCropTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Filename"
   * </p>
   */

  @DISPID(1415)
  @PropGet
  java.lang.String getFilename();


  /**
   * <p>
   * Setter method for the COM property "Filename"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1415)
  @PropPut
  void setFilename(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  float getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "LockAspectRatio"
   * </p>
   */

  @DISPID(1700)
  @PropGet
  com.exceljava.com4j.office.MsoTriState getLockAspectRatio();


  /**
   * <p>
   * Setter method for the COM property "LockAspectRatio"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(1700)
  @PropPut
  void setLockAspectRatio(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  float getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    float rhs);


  // Properties:
}
