package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSharedWorkspaceTaskPriority implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSharedWorkspaceTaskPriorityHigh(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSharedWorkspaceTaskPriorityNormal(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSharedWorkspaceTaskPriorityLow(3),
  ;

  private final int value;
  MsoSharedWorkspaceTaskPriority(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
