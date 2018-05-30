package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetTextEffectShape implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextEffectShapeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTextEffectShapePlainText(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextEffectShapeStop(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTextEffectShapeTriangleUp(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTextEffectShapeTriangleDown(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoTextEffectShapeChevronUp(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoTextEffectShapeChevronDown(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoTextEffectShapeRingInside(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoTextEffectShapeRingOutside(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoTextEffectShapeArchUpCurve(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoTextEffectShapeArchDownCurve(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoTextEffectShapeCircleCurve(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoTextEffectShapeButtonCurve(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoTextEffectShapeArchUpPour(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoTextEffectShapeArchDownPour(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoTextEffectShapeCirclePour(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoTextEffectShapeButtonPour(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoTextEffectShapeCurveUp(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoTextEffectShapeCurveDown(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoTextEffectShapeCanUp(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoTextEffectShapeCanDown(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoTextEffectShapeWave1(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoTextEffectShapeWave2(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoTextEffectShapeDoubleWave1(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoTextEffectShapeDoubleWave2(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoTextEffectShapeInflate(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoTextEffectShapeDeflate(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  msoTextEffectShapeInflateBottom(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  msoTextEffectShapeDeflateBottom(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  msoTextEffectShapeInflateTop(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  msoTextEffectShapeDeflateTop(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  msoTextEffectShapeDeflateInflate(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  msoTextEffectShapeDeflateInflateDeflate(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  msoTextEffectShapeFadeRight(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  msoTextEffectShapeFadeLeft(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  msoTextEffectShapeFadeUp(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  msoTextEffectShapeFadeDown(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  msoTextEffectShapeSlantUp(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  msoTextEffectShapeSlantDown(38),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  msoTextEffectShapeCascadeUp(39),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  msoTextEffectShapeCascadeDown(40),
  ;

  private final int value;
  MsoPresetTextEffectShape(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
