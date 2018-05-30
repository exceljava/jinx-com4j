package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPictureColorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPictureMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoPictureAutomatic(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoPictureGrayscale(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoPictureBlackAndWhite(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoPictureWatermark(4),
  ;

  private final int value;
  MsoPictureColorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
