package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCopyPictureFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlBitmap(2),
  /**
   * <p>
   * The value of this constant is -4147
   * </p>
   */
  xlPicture(-4147),
  ;

  private final int value;
  XlCopyPictureFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
