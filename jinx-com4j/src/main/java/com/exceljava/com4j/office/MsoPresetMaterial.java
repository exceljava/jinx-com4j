package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetMaterial implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPresetMaterialMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoMaterialMatte(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoMaterialPlastic(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoMaterialMetal(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoMaterialWireFrame(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoMaterialMatte2(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoMaterialPlastic2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoMaterialMetal2(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoMaterialWarmMatte(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoMaterialTranslucentPowder(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoMaterialPowder(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoMaterialDarkEdge(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoMaterialSoftEdge(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoMaterialClear(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoMaterialFlat(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoMaterialSoftMetal(15),
  ;

  private final int value;
  MsoPresetMaterial(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
