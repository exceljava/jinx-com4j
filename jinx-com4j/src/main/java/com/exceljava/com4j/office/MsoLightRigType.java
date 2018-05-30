package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLightRigType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLightRigMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLightRigLegacyFlat1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLightRigLegacyFlat2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLightRigLegacyFlat3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLightRigLegacyFlat4(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLightRigLegacyNormal1(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoLightRigLegacyNormal2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoLightRigLegacyNormal3(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoLightRigLegacyNormal4(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoLightRigLegacyHarsh1(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoLightRigLegacyHarsh2(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoLightRigLegacyHarsh3(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoLightRigLegacyHarsh4(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoLightRigThreePoint(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoLightRigBalanced(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoLightRigSoft(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoLightRigHarsh(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoLightRigFlood(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoLightRigContrasting(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoLightRigMorning(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoLightRigSunrise(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoLightRigSunset(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoLightRigChilly(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoLightRigFreezing(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoLightRigFlat(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoLightRigTwoPoint(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoLightRigGlow(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  msoLightRigBrightRoom(27),
  ;

  private final int value;
  MsoLightRigType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
