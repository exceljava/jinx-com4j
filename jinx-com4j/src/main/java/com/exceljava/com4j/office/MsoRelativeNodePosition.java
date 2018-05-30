package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoRelativeNodePosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBeforeNode(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAfterNode(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBeforeFirstSibling(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoAfterLastSibling(4),
  ;

  private final int value;
  MsoRelativeNodePosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
