package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileFindView implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoViewFileInfo(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoViewPreview(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoViewSummaryInfo(3),
  ;

  private final int value;
  MsoFileFindView(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
