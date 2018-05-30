package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBlogImageType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoblogImageTypeJPEG(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoblogImageTypeGIF(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoblogImageTypePNG(3),
  ;

  private final int value;
  MsoBlogImageType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
