package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoMergeCmd implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoMergeUnion(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoMergeCombine(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoMergeIntersect(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoMergeSubtract(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoMergeFragment(5),
  ;

  private final int value;
  MsoMergeCmd(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
