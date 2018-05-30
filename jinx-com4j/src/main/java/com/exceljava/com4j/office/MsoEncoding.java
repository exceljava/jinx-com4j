package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoEncoding implements ComEnum {
  /**
   * <p>
   * The value of this constant is 874
   * </p>
   */
  msoEncodingThai(874),
  /**
   * <p>
   * The value of this constant is 932
   * </p>
   */
  msoEncodingJapaneseShiftJIS(932),
  /**
   * <p>
   * The value of this constant is 936
   * </p>
   */
  msoEncodingSimplifiedChineseGBK(936),
  /**
   * <p>
   * The value of this constant is 949
   * </p>
   */
  msoEncodingKorean(949),
  /**
   * <p>
   * The value of this constant is 950
   * </p>
   */
  msoEncodingTraditionalChineseBig5(950),
  /**
   * <p>
   * The value of this constant is 1200
   * </p>
   */
  msoEncodingUnicodeLittleEndian(1200),
  /**
   * <p>
   * The value of this constant is 1201
   * </p>
   */
  msoEncodingUnicodeBigEndian(1201),
  /**
   * <p>
   * The value of this constant is 1250
   * </p>
   */
  msoEncodingCentralEuropean(1250),
  /**
   * <p>
   * The value of this constant is 1251
   * </p>
   */
  msoEncodingCyrillic(1251),
  /**
   * <p>
   * The value of this constant is 1252
   * </p>
   */
  msoEncodingWestern(1252),
  /**
   * <p>
   * The value of this constant is 1253
   * </p>
   */
  msoEncodingGreek(1253),
  /**
   * <p>
   * The value of this constant is 1254
   * </p>
   */
  msoEncodingTurkish(1254),
  /**
   * <p>
   * The value of this constant is 1255
   * </p>
   */
  msoEncodingHebrew(1255),
  /**
   * <p>
   * The value of this constant is 1256
   * </p>
   */
  msoEncodingArabic(1256),
  /**
   * <p>
   * The value of this constant is 1257
   * </p>
   */
  msoEncodingBaltic(1257),
  /**
   * <p>
   * The value of this constant is 1258
   * </p>
   */
  msoEncodingVietnamese(1258),
  /**
   * <p>
   * The value of this constant is 50001
   * </p>
   */
  msoEncodingAutoDetect(50001),
  /**
   * <p>
   * The value of this constant is 50932
   * </p>
   */
  msoEncodingJapaneseAutoDetect(50932),
  /**
   * <p>
   * The value of this constant is 50936
   * </p>
   */
  msoEncodingSimplifiedChineseAutoDetect(50936),
  /**
   * <p>
   * The value of this constant is 50949
   * </p>
   */
  msoEncodingKoreanAutoDetect(50949),
  /**
   * <p>
   * The value of this constant is 50950
   * </p>
   */
  msoEncodingTraditionalChineseAutoDetect(50950),
  /**
   * <p>
   * The value of this constant is 51251
   * </p>
   */
  msoEncodingCyrillicAutoDetect(51251),
  /**
   * <p>
   * The value of this constant is 51253
   * </p>
   */
  msoEncodingGreekAutoDetect(51253),
  /**
   * <p>
   * The value of this constant is 51256
   * </p>
   */
  msoEncodingArabicAutoDetect(51256),
  /**
   * <p>
   * The value of this constant is 28591
   * </p>
   */
  msoEncodingISO88591Latin1(28591),
  /**
   * <p>
   * The value of this constant is 28592
   * </p>
   */
  msoEncodingISO88592CentralEurope(28592),
  /**
   * <p>
   * The value of this constant is 28593
   * </p>
   */
  msoEncodingISO88593Latin3(28593),
  /**
   * <p>
   * The value of this constant is 28594
   * </p>
   */
  msoEncodingISO88594Baltic(28594),
  /**
   * <p>
   * The value of this constant is 28595
   * </p>
   */
  msoEncodingISO88595Cyrillic(28595),
  /**
   * <p>
   * The value of this constant is 28596
   * </p>
   */
  msoEncodingISO88596Arabic(28596),
  /**
   * <p>
   * The value of this constant is 28597
   * </p>
   */
  msoEncodingISO88597Greek(28597),
  /**
   * <p>
   * The value of this constant is 28598
   * </p>
   */
  msoEncodingISO88598Hebrew(28598),
  /**
   * <p>
   * The value of this constant is 28599
   * </p>
   */
  msoEncodingISO88599Turkish(28599),
  /**
   * <p>
   * The value of this constant is 28605
   * </p>
   */
  msoEncodingISO885915Latin9(28605),
  /**
   * <p>
   * The value of this constant is 38598
   * </p>
   */
  msoEncodingISO88598HebrewLogical(38598),
  /**
   * <p>
   * The value of this constant is 50220
   * </p>
   */
  msoEncodingISO2022JPNoHalfwidthKatakana(50220),
  /**
   * <p>
   * The value of this constant is 50221
   * </p>
   */
  msoEncodingISO2022JPJISX02021984(50221),
  /**
   * <p>
   * The value of this constant is 50222
   * </p>
   */
  msoEncodingISO2022JPJISX02011989(50222),
  /**
   * <p>
   * The value of this constant is 50225
   * </p>
   */
  msoEncodingISO2022KR(50225),
  /**
   * <p>
   * The value of this constant is 50227
   * </p>
   */
  msoEncodingISO2022CNTraditionalChinese(50227),
  /**
   * <p>
   * The value of this constant is 50229
   * </p>
   */
  msoEncodingISO2022CNSimplifiedChinese(50229),
  /**
   * <p>
   * The value of this constant is 10000
   * </p>
   */
  msoEncodingMacRoman(10000),
  /**
   * <p>
   * The value of this constant is 10001
   * </p>
   */
  msoEncodingMacJapanese(10001),
  /**
   * <p>
   * The value of this constant is 10002
   * </p>
   */
  msoEncodingMacTraditionalChineseBig5(10002),
  /**
   * <p>
   * The value of this constant is 10003
   * </p>
   */
  msoEncodingMacKorean(10003),
  /**
   * <p>
   * The value of this constant is 10004
   * </p>
   */
  msoEncodingMacArabic(10004),
  /**
   * <p>
   * The value of this constant is 10005
   * </p>
   */
  msoEncodingMacHebrew(10005),
  /**
   * <p>
   * The value of this constant is 10006
   * </p>
   */
  msoEncodingMacGreek1(10006),
  /**
   * <p>
   * The value of this constant is 10007
   * </p>
   */
  msoEncodingMacCyrillic(10007),
  /**
   * <p>
   * The value of this constant is 10008
   * </p>
   */
  msoEncodingMacSimplifiedChineseGB2312(10008),
  /**
   * <p>
   * The value of this constant is 10010
   * </p>
   */
  msoEncodingMacRomania(10010),
  /**
   * <p>
   * The value of this constant is 10017
   * </p>
   */
  msoEncodingMacUkraine(10017),
  /**
   * <p>
   * The value of this constant is 10029
   * </p>
   */
  msoEncodingMacLatin2(10029),
  /**
   * <p>
   * The value of this constant is 10079
   * </p>
   */
  msoEncodingMacIcelandic(10079),
  /**
   * <p>
   * The value of this constant is 10081
   * </p>
   */
  msoEncodingMacTurkish(10081),
  /**
   * <p>
   * The value of this constant is 10082
   * </p>
   */
  msoEncodingMacCroatia(10082),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  msoEncodingEBCDICUSCanada(37),
  /**
   * <p>
   * The value of this constant is 500
   * </p>
   */
  msoEncodingEBCDICInternational(500),
  /**
   * <p>
   * The value of this constant is 870
   * </p>
   */
  msoEncodingEBCDICMultilingualROECELatin2(870),
  /**
   * <p>
   * The value of this constant is 875
   * </p>
   */
  msoEncodingEBCDICGreekModern(875),
  /**
   * <p>
   * The value of this constant is 1026
   * </p>
   */
  msoEncodingEBCDICTurkishLatin5(1026),
  /**
   * <p>
   * The value of this constant is 20273
   * </p>
   */
  msoEncodingEBCDICGermany(20273),
  /**
   * <p>
   * The value of this constant is 20277
   * </p>
   */
  msoEncodingEBCDICDenmarkNorway(20277),
  /**
   * <p>
   * The value of this constant is 20278
   * </p>
   */
  msoEncodingEBCDICFinlandSweden(20278),
  /**
   * <p>
   * The value of this constant is 20280
   * </p>
   */
  msoEncodingEBCDICItaly(20280),
  /**
   * <p>
   * The value of this constant is 20284
   * </p>
   */
  msoEncodingEBCDICLatinAmericaSpain(20284),
  /**
   * <p>
   * The value of this constant is 20285
   * </p>
   */
  msoEncodingEBCDICUnitedKingdom(20285),
  /**
   * <p>
   * The value of this constant is 20290
   * </p>
   */
  msoEncodingEBCDICJapaneseKatakanaExtended(20290),
  /**
   * <p>
   * The value of this constant is 20297
   * </p>
   */
  msoEncodingEBCDICFrance(20297),
  /**
   * <p>
   * The value of this constant is 20420
   * </p>
   */
  msoEncodingEBCDICArabic(20420),
  /**
   * <p>
   * The value of this constant is 20423
   * </p>
   */
  msoEncodingEBCDICGreek(20423),
  /**
   * <p>
   * The value of this constant is 20424
   * </p>
   */
  msoEncodingEBCDICHebrew(20424),
  /**
   * <p>
   * The value of this constant is 20833
   * </p>
   */
  msoEncodingEBCDICKoreanExtended(20833),
  /**
   * <p>
   * The value of this constant is 20838
   * </p>
   */
  msoEncodingEBCDICThai(20838),
  /**
   * <p>
   * The value of this constant is 20871
   * </p>
   */
  msoEncodingEBCDICIcelandic(20871),
  /**
   * <p>
   * The value of this constant is 20905
   * </p>
   */
  msoEncodingEBCDICTurkish(20905),
  /**
   * <p>
   * The value of this constant is 20880
   * </p>
   */
  msoEncodingEBCDICRussian(20880),
  /**
   * <p>
   * The value of this constant is 21025
   * </p>
   */
  msoEncodingEBCDICSerbianBulgarian(21025),
  /**
   * <p>
   * The value of this constant is 50930
   * </p>
   */
  msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese(50930),
  /**
   * <p>
   * The value of this constant is 50931
   * </p>
   */
  msoEncodingEBCDICUSCanadaAndJapanese(50931),
  /**
   * <p>
   * The value of this constant is 50933
   * </p>
   */
  msoEncodingEBCDICKoreanExtendedAndKorean(50933),
  /**
   * <p>
   * The value of this constant is 50935
   * </p>
   */
  msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese(50935),
  /**
   * <p>
   * The value of this constant is 50937
   * </p>
   */
  msoEncodingEBCDICUSCanadaAndTraditionalChinese(50937),
  /**
   * <p>
   * The value of this constant is 50939
   * </p>
   */
  msoEncodingEBCDICJapaneseLatinExtendedAndJapanese(50939),
  /**
   * <p>
   * The value of this constant is 437
   * </p>
   */
  msoEncodingOEMUnitedStates(437),
  /**
   * <p>
   * The value of this constant is 737
   * </p>
   */
  msoEncodingOEMGreek437G(737),
  /**
   * <p>
   * The value of this constant is 775
   * </p>
   */
  msoEncodingOEMBaltic(775),
  /**
   * <p>
   * The value of this constant is 850
   * </p>
   */
  msoEncodingOEMMultilingualLatinI(850),
  /**
   * <p>
   * The value of this constant is 852
   * </p>
   */
  msoEncodingOEMMultilingualLatinII(852),
  /**
   * <p>
   * The value of this constant is 855
   * </p>
   */
  msoEncodingOEMCyrillic(855),
  /**
   * <p>
   * The value of this constant is 857
   * </p>
   */
  msoEncodingOEMTurkish(857),
  /**
   * <p>
   * The value of this constant is 860
   * </p>
   */
  msoEncodingOEMPortuguese(860),
  /**
   * <p>
   * The value of this constant is 861
   * </p>
   */
  msoEncodingOEMIcelandic(861),
  /**
   * <p>
   * The value of this constant is 862
   * </p>
   */
  msoEncodingOEMHebrew(862),
  /**
   * <p>
   * The value of this constant is 863
   * </p>
   */
  msoEncodingOEMCanadianFrench(863),
  /**
   * <p>
   * The value of this constant is 864
   * </p>
   */
  msoEncodingOEMArabic(864),
  /**
   * <p>
   * The value of this constant is 865
   * </p>
   */
  msoEncodingOEMNordic(865),
  /**
   * <p>
   * The value of this constant is 866
   * </p>
   */
  msoEncodingOEMCyrillicII(866),
  /**
   * <p>
   * The value of this constant is 869
   * </p>
   */
  msoEncodingOEMModernGreek(869),
  /**
   * <p>
   * The value of this constant is 51932
   * </p>
   */
  msoEncodingEUCJapanese(51932),
  /**
   * <p>
   * The value of this constant is 51936
   * </p>
   */
  msoEncodingEUCChineseSimplifiedChinese(51936),
  /**
   * <p>
   * The value of this constant is 51949
   * </p>
   */
  msoEncodingEUCKorean(51949),
  /**
   * <p>
   * The value of this constant is 51950
   * </p>
   */
  msoEncodingEUCTaiwaneseTraditionalChinese(51950),
  /**
   * <p>
   * The value of this constant is 57002
   * </p>
   */
  msoEncodingISCIIDevanagari(57002),
  /**
   * <p>
   * The value of this constant is 57003
   * </p>
   */
  msoEncodingISCIIBengali(57003),
  /**
   * <p>
   * The value of this constant is 57004
   * </p>
   */
  msoEncodingISCIITamil(57004),
  /**
   * <p>
   * The value of this constant is 57005
   * </p>
   */
  msoEncodingISCIITelugu(57005),
  /**
   * <p>
   * The value of this constant is 57006
   * </p>
   */
  msoEncodingISCIIAssamese(57006),
  /**
   * <p>
   * The value of this constant is 57007
   * </p>
   */
  msoEncodingISCIIOriya(57007),
  /**
   * <p>
   * The value of this constant is 57008
   * </p>
   */
  msoEncodingISCIIKannada(57008),
  /**
   * <p>
   * The value of this constant is 57009
   * </p>
   */
  msoEncodingISCIIMalayalam(57009),
  /**
   * <p>
   * The value of this constant is 57010
   * </p>
   */
  msoEncodingISCIIGujarati(57010),
  /**
   * <p>
   * The value of this constant is 57011
   * </p>
   */
  msoEncodingISCIIPunjabi(57011),
  /**
   * <p>
   * The value of this constant is 708
   * </p>
   */
  msoEncodingArabicASMO(708),
  /**
   * <p>
   * The value of this constant is 720
   * </p>
   */
  msoEncodingArabicTransparentASMO(720),
  /**
   * <p>
   * The value of this constant is 1361
   * </p>
   */
  msoEncodingKoreanJohab(1361),
  /**
   * <p>
   * The value of this constant is 20000
   * </p>
   */
  msoEncodingTaiwanCNS(20000),
  /**
   * <p>
   * The value of this constant is 20001
   * </p>
   */
  msoEncodingTaiwanTCA(20001),
  /**
   * <p>
   * The value of this constant is 20002
   * </p>
   */
  msoEncodingTaiwanEten(20002),
  /**
   * <p>
   * The value of this constant is 20003
   * </p>
   */
  msoEncodingTaiwanIBM5550(20003),
  /**
   * <p>
   * The value of this constant is 20004
   * </p>
   */
  msoEncodingTaiwanTeleText(20004),
  /**
   * <p>
   * The value of this constant is 20005
   * </p>
   */
  msoEncodingTaiwanWang(20005),
  /**
   * <p>
   * The value of this constant is 20105
   * </p>
   */
  msoEncodingIA5IRV(20105),
  /**
   * <p>
   * The value of this constant is 20106
   * </p>
   */
  msoEncodingIA5German(20106),
  /**
   * <p>
   * The value of this constant is 20107
   * </p>
   */
  msoEncodingIA5Swedish(20107),
  /**
   * <p>
   * The value of this constant is 20108
   * </p>
   */
  msoEncodingIA5Norwegian(20108),
  /**
   * <p>
   * The value of this constant is 20127
   * </p>
   */
  msoEncodingUSASCII(20127),
  /**
   * <p>
   * The value of this constant is 20261
   * </p>
   */
  msoEncodingT61(20261),
  /**
   * <p>
   * The value of this constant is 20269
   * </p>
   */
  msoEncodingISO6937NonSpacingAccent(20269),
  /**
   * <p>
   * The value of this constant is 20866
   * </p>
   */
  msoEncodingKOI8R(20866),
  /**
   * <p>
   * The value of this constant is 21027
   * </p>
   */
  msoEncodingExtAlphaLowercase(21027),
  /**
   * <p>
   * The value of this constant is 21866
   * </p>
   */
  msoEncodingKOI8U(21866),
  /**
   * <p>
   * The value of this constant is 29001
   * </p>
   */
  msoEncodingEuropa3(29001),
  /**
   * <p>
   * The value of this constant is 52936
   * </p>
   */
  msoEncodingHZGBSimplifiedChinese(52936),
  /**
   * <p>
   * The value of this constant is 54936
   * </p>
   */
  msoEncodingSimplifiedChineseGB18030(54936),
  /**
   * <p>
   * The value of this constant is 65000
   * </p>
   */
  msoEncodingUTF7(65000),
  /**
   * <p>
   * The value of this constant is 65001
   * </p>
   */
  msoEncodingUTF8(65001),
  ;

  private final int value;
  MsoEncoding(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
