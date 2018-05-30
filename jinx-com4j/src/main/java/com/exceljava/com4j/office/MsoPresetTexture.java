package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetTexture implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPresetTextureMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTexturePapyrus(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextureCanvas(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTextureDenim(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTextureWovenMat(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoTextureWaterDroplets(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoTexturePaperBag(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoTextureFishFossil(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoTextureSand(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoTextureGreenMarble(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoTextureWhiteMarble(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoTextureBrownMarble(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoTextureGranite(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoTextureNewsprint(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoTextureRecycledPaper(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoTextureParchment(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoTextureStationery(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoTextureBlueTissuePaper(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoTexturePinkTissuePaper(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoTexturePurpleMesh(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoTextureBouquet(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoTextureCork(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoTextureWalnut(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoTextureOak(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoTextureMediumWood(24),
  ;

  private final int value;
  MsoPresetTexture(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
