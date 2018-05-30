package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoNumberedBulletStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBulletStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoBulletAlphaLCPeriod(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBulletAlphaUCPeriod(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBulletArabicParenRight(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBulletArabicPeriod(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoBulletRomanLCParenBoth(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoBulletRomanLCParenRight(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoBulletRomanLCPeriod(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoBulletRomanUCPeriod(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoBulletAlphaLCParenBoth(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoBulletAlphaLCParenRight(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoBulletAlphaUCParenBoth(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoBulletAlphaUCParenRight(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoBulletArabicParenBoth(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoBulletArabicPlain(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoBulletRomanUCParenBoth(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoBulletRomanUCParenRight(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoBulletSimpChinPlain(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoBulletSimpChinPeriod(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoBulletCircleNumDBPlain(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoBulletCircleNumWDWhitePlain(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoBulletCircleNumWDBlackPlain(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoBulletTradChinPlain(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoBulletTradChinPeriod(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoBulletArabicAlphaDash(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoBulletArabicAbjadDash(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoBulletHebrewAlphaDash(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoBulletKanjiKoreanPlain(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  msoBulletKanjiKoreanPeriod(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  msoBulletArabicDBPlain(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  msoBulletArabicDBPeriod(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  msoBulletThaiAlphaPeriod(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  msoBulletThaiAlphaParenRight(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  msoBulletThaiAlphaParenBoth(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  msoBulletThaiNumPeriod(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  msoBulletThaiNumParenRight(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  msoBulletThaiNumParenBoth(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  msoBulletHindiAlphaPeriod(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  msoBulletHindiNumPeriod(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  msoBulletKanjiSimpChinDBPeriod(38),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  msoBulletHindiNumParenRight(39),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  msoBulletHindiAlpha1Period(40),
  ;

  private final int value;
  MsoNumberedBulletStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
