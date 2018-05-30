package com.exceljava.com4j.vbide  ;

import com4j.*;

/**
 */
public enum vbext_ProjectType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 100
   * </p>
   */
  vbext_pt_HostProject(100),
  /**
   * <p>
   * The value of this constant is 101
   * </p>
   */
  vbext_pt_StandAlone(101),
  ;

  private final int value;
  vbext_ProjectType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
