package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPictureType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPictureTypeDefault(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoPictureTypePNG(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoPictureTypeBMP(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoPictureTypeGIF(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoPictureTypeJPG(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoPictureTypePDF(4),
  ;

  private final int value;
  MsoPictureType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
