package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLocationInTable implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4110
   * </p>
   */
  xlColumnHeader(-4110),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlColumnItem(5),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlDataHeader(3),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlDataItem(7),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPageHeader(2),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlPageItem(6),
  /**
   * <p>
   * The value of this constant is -4153
   * </p>
   */
  xlRowHeader(-4153),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlRowItem(4),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlTableBody(8),
  ;

  private final int value;
  XlLocationInTable(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
