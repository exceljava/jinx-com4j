package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSharedWorkspaceTaskStatus implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSharedWorkspaceTaskStatusNotStarted(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSharedWorkspaceTaskStatusInProgress(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSharedWorkspaceTaskStatusCompleted(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoSharedWorkspaceTaskStatusDeferred(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoSharedWorkspaceTaskStatusWaiting(5),
  ;

  private final int value;
  MsoSharedWorkspaceTaskStatus(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
