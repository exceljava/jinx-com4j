package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSortBy implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSortByFileName(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSortBySize(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSortByFileType(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoSortByLastModified(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoSortByNone(5),
  ;

  private final int value;
  MsoSortBy(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
