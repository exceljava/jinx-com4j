package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SeriesGradientStopColorFormat extends Com4jObject {
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
   * Getter method for the COM property "TintAndShade"
   * </p>
   */

  @DISPID(2366)
  @PropGet
  float getTintAndShade();


  /**
   * <p>
   * Setter method for the COM property "TintAndShade"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(2366)
  @PropPut
  void setTintAndShade(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ObjectThemeColor"
   * </p>
   */

  @DISPID(3265)
  @PropGet
  com.exceljava.com4j.office.MsoThemeColorIndex getObjectThemeColor();


  /**
   * <p>
   * Setter method for the COM property "ObjectThemeColor"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoThemeColorIndex parameter.
   */

  @DISPID(3265)
  @PropPut
  void setObjectThemeColor(
    com.exceljava.com4j.office.MsoThemeColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "RGB"
   * </p>
   */

  @DISPID(1055)
  @PropGet
  int getRGB();


  /**
   * <p>
   * Setter method for the COM property "RGB"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1055)
  @PropPut
  void setRGB(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Transparency"
   * </p>
   */

  @DISPID(3266)
  @PropGet
  float getTransparency();


  /**
   * <p>
   * Setter method for the COM property "Transparency"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @DISPID(3266)
  @PropPut
  void setTransparency(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.office.MsoColorType getType();


  // Properties:
}
