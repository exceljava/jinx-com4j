package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPictureCompress implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoPictureCompressDocDefault(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoPictureCompressFalse(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoPictureCompressTrue(1),
  ;

  private final int value;
  MsoPictureCompress(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
