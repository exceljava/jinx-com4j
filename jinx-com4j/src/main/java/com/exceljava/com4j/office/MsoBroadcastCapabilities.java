package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBroadcastCapabilities implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  BroadcastCapFileSizeLimited(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  BroadcastCapSupportsMeetingNotes(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  BroadcastCapSupportsUpdateDoc(4),
  ;

  private final int value;
  MsoBroadcastCapabilities(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
