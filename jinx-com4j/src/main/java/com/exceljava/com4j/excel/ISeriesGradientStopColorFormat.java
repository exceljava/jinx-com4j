package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244F9-0001-0000-C000-000000000046}")
public interface ISeriesGradientStopColorFormat extends Com4jObject {
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
   * Getter method for the COM property "TintAndShade"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(10)
  float getTintAndShade();


  /**
   * <p>
   * Setter method for the COM property "TintAndShade"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(11)
  void setTintAndShade(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ObjectThemeColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoThemeColorIndex
   */

  @VTID(12)
  com.exceljava.com4j.office.MsoThemeColorIndex getObjectThemeColor();


  /**
   * <p>
   * Setter method for the COM property "ObjectThemeColor"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoThemeColorIndex parameter.
   */

  @VTID(13)
  void setObjectThemeColor(
    com.exceljava.com4j.office.MsoThemeColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "RGB"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
  int getRGB();


  /**
   * <p>
   * Setter method for the COM property "RGB"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(15)
  void setRGB(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Transparency"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(16)
  float getTransparency();


  /**
   * <p>
   * Setter method for the COM property "Transparency"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(17)
  void setTransparency(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoColorType
   */

  @VTID(18)
  com.exceljava.com4j.office.MsoColorType getType();


  // Properties:
}
