package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSyncStatusType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoSyncStatusNoSharedWorkspace(0),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoSyncStatusNotRoaming(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSyncStatusLatest(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSyncStatusNewerAvailable(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSyncStatusLocalChanges(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoSyncStatusConflict(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoSyncStatusSuspended(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoSyncStatusError(6),
  ;

  private final int value;
  MsoSyncStatusType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
