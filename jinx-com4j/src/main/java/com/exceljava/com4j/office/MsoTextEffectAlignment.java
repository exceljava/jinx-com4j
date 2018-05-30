package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextEffectAlignment implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextEffectAlignmentMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTextEffectAlignmentLeft(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextEffectAlignmentCentered(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTextEffectAlignmentRight(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTextEffectAlignmentLetterJustify(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoTextEffectAlignmentWordJustify(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoTextEffectAlignmentStretchJustify(6),
  ;

  private final int value;
  MsoTextEffectAlignment(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
