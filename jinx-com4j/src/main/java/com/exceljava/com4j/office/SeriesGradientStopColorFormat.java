package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{08D6CD0F-98AA-468B-81F3-A6B2CB6C84C9}")
public interface SeriesGradientStopColorFormat extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "TintAndShade"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(8)
  float getTintAndShade();


  /**
   * <p>
   * Setter method for the COM property "TintAndShade"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(9)
  void setTintAndShade(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ObjectThemeColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoThemeColorIndex
   */

  @VTID(10)
  com.exceljava.com4j.office.MsoThemeColorIndex getObjectThemeColor();


  /**
   * <p>
   * Setter method for the COM property "ObjectThemeColor"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoThemeColorIndex parameter.
   */

  @VTID(11)
  void setObjectThemeColor(
    com.exceljava.com4j.office.MsoThemeColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "RGB"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getRGB();


  /**
   * <p>
   * Setter method for the COM property "RGB"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(13)
  void setRGB(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Transparency"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(14)
  float getTransparency();


  /**
   * <p>
   * Setter method for the COM property "Transparency"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(15)
  void setTransparency(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory Holder<com.exceljava.com4j.office.MsoColorType> parameter.
   */

  @VTID(16)
  void getType(
    Holder<com.exceljava.com4j.office.MsoColorType> rhs);


  // Properties:
}
