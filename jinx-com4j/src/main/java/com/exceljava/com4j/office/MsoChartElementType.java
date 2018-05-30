package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoChartElementType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoElementChartTitleNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoElementChartTitleCenteredOverlay(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoElementChartTitleAboveChart(2),
  /**
   * <p>
   * The value of this constant is 100
   * </p>
   */
  msoElementLegendNone(100),
  /**
   * <p>
   * The value of this constant is 101
   * </p>
   */
  msoElementLegendRight(101),
  /**
   * <p>
   * The value of this constant is 102
   * </p>
   */
  msoElementLegendTop(102),
  /**
   * <p>
   * The value of this constant is 103
   * </p>
   */
  msoElementLegendLeft(103),
  /**
   * <p>
   * The value of this constant is 104
   * </p>
   */
  msoElementLegendBottom(104),
  /**
   * <p>
   * The value of this constant is 105
   * </p>
   */
  msoElementLegendRightOverlay(105),
  /**
   * <p>
   * The value of this constant is 106
   * </p>
   */
  msoElementLegendLeftOverlay(106),
  /**
   * <p>
   * The value of this constant is 200
   * </p>
   */
  msoElementDataLabelNone(200),
  /**
   * <p>
   * The value of this constant is 201
   * </p>
   */
  msoElementDataLabelShow(201),
  /**
   * <p>
   * The value of this constant is 202
   * </p>
   */
  msoElementDataLabelCenter(202),
  /**
   * <p>
   * The value of this constant is 203
   * </p>
   */
  msoElementDataLabelInsideEnd(203),
  /**
   * <p>
   * The value of this constant is 204
   * </p>
   */
  msoElementDataLabelInsideBase(204),
  /**
   * <p>
   * The value of this constant is 205
   * </p>
   */
  msoElementDataLabelOutSideEnd(205),
  /**
   * <p>
   * The value of this constant is 206
   * </p>
   */
  msoElementDataLabelLeft(206),
  /**
   * <p>
   * The value of this constant is 207
   * </p>
   */
  msoElementDataLabelRight(207),
  /**
   * <p>
   * The value of this constant is 208
   * </p>
   */
  msoElementDataLabelTop(208),
  /**
   * <p>
   * The value of this constant is 209
   * </p>
   */
  msoElementDataLabelBottom(209),
  /**
   * <p>
   * The value of this constant is 210
   * </p>
   */
  msoElementDataLabelBestFit(210),
  /**
   * <p>
   * The value of this constant is 211
   * </p>
   */
  msoElementDataLabelCallout(211),
  /**
   * <p>
   * The value of this constant is 300
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleNone(300),
  /**
   * <p>
   * The value of this constant is 301
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleAdjacentToAxis(301),
  /**
   * <p>
   * The value of this constant is 302
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleBelowAxis(302),
  /**
   * <p>
   * The value of this constant is 303
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleRotated(303),
  /**
   * <p>
   * The value of this constant is 304
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleVertical(304),
  /**
   * <p>
   * The value of this constant is 305
   * </p>
   */
  msoElementPrimaryCategoryAxisTitleHorizontal(305),
  /**
   * <p>
   * The value of this constant is 306
   * </p>
   */
  msoElementPrimaryValueAxisTitleNone(306),
  /**
   * <p>
   * The value of this constant is 306
   * </p>
   */
  msoElementPrimaryValueAxisTitleAdjacentToAxis(306),
  /**
   * <p>
   * The value of this constant is 308
   * </p>
   */
  msoElementPrimaryValueAxisTitleBelowAxis(308),
  /**
   * <p>
   * The value of this constant is 309
   * </p>
   */
  msoElementPrimaryValueAxisTitleRotated(309),
  /**
   * <p>
   * The value of this constant is 310
   * </p>
   */
  msoElementPrimaryValueAxisTitleVertical(310),
  /**
   * <p>
   * The value of this constant is 311
   * </p>
   */
  msoElementPrimaryValueAxisTitleHorizontal(311),
  /**
   * <p>
   * The value of this constant is 312
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleNone(312),
  /**
   * <p>
   * The value of this constant is 313
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleAdjacentToAxis(313),
  /**
   * <p>
   * The value of this constant is 314
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleBelowAxis(314),
  /**
   * <p>
   * The value of this constant is 315
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleRotated(315),
  /**
   * <p>
   * The value of this constant is 316
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleVertical(316),
  /**
   * <p>
   * The value of this constant is 317
   * </p>
   */
  msoElementSecondaryCategoryAxisTitleHorizontal(317),
  /**
   * <p>
   * The value of this constant is 318
   * </p>
   */
  msoElementSecondaryValueAxisTitleNone(318),
  /**
   * <p>
   * The value of this constant is 319
   * </p>
   */
  msoElementSecondaryValueAxisTitleAdjacentToAxis(319),
  /**
   * <p>
   * The value of this constant is 320
   * </p>
   */
  msoElementSecondaryValueAxisTitleBelowAxis(320),
  /**
   * <p>
   * The value of this constant is 321
   * </p>
   */
  msoElementSecondaryValueAxisTitleRotated(321),
  /**
   * <p>
   * The value of this constant is 322
   * </p>
   */
  msoElementSecondaryValueAxisTitleVertical(322),
  /**
   * <p>
   * The value of this constant is 323
   * </p>
   */
  msoElementSecondaryValueAxisTitleHorizontal(323),
  /**
   * <p>
   * The value of this constant is 324
   * </p>
   */
  msoElementSeriesAxisTitleNone(324),
  /**
   * <p>
   * The value of this constant is 325
   * </p>
   */
  msoElementSeriesAxisTitleRotated(325),
  /**
   * <p>
   * The value of this constant is 326
   * </p>
   */
  msoElementSeriesAxisTitleVertical(326),
  /**
   * <p>
   * The value of this constant is 327
   * </p>
   */
  msoElementSeriesAxisTitleHorizontal(327),
  /**
   * <p>
   * The value of this constant is 328
   * </p>
   */
  msoElementPrimaryValueGridLinesNone(328),
  /**
   * <p>
   * The value of this constant is 329
   * </p>
   */
  msoElementPrimaryValueGridLinesMinor(329),
  /**
   * <p>
   * The value of this constant is 330
   * </p>
   */
  msoElementPrimaryValueGridLinesMajor(330),
  /**
   * <p>
   * The value of this constant is 331
   * </p>
   */
  msoElementPrimaryValueGridLinesMinorMajor(331),
  /**
   * <p>
   * The value of this constant is 332
   * </p>
   */
  msoElementPrimaryCategoryGridLinesNone(332),
  /**
   * <p>
   * The value of this constant is 333
   * </p>
   */
  msoElementPrimaryCategoryGridLinesMinor(333),
  /**
   * <p>
   * The value of this constant is 334
   * </p>
   */
  msoElementPrimaryCategoryGridLinesMajor(334),
  /**
   * <p>
   * The value of this constant is 335
   * </p>
   */
  msoElementPrimaryCategoryGridLinesMinorMajor(335),
  /**
   * <p>
   * The value of this constant is 336
   * </p>
   */
  msoElementSecondaryValueGridLinesNone(336),
  /**
   * <p>
   * The value of this constant is 337
   * </p>
   */
  msoElementSecondaryValueGridLinesMinor(337),
  /**
   * <p>
   * The value of this constant is 338
   * </p>
   */
  msoElementSecondaryValueGridLinesMajor(338),
  /**
   * <p>
   * The value of this constant is 339
   * </p>
   */
  msoElementSecondaryValueGridLinesMinorMajor(339),
  /**
   * <p>
   * The value of this constant is 340
   * </p>
   */
  msoElementSecondaryCategoryGridLinesNone(340),
  /**
   * <p>
   * The value of this constant is 341
   * </p>
   */
  msoElementSecondaryCategoryGridLinesMinor(341),
  /**
   * <p>
   * The value of this constant is 342
   * </p>
   */
  msoElementSecondaryCategoryGridLinesMajor(342),
  /**
   * <p>
   * The value of this constant is 343
   * </p>
   */
  msoElementSecondaryCategoryGridLinesMinorMajor(343),
  /**
   * <p>
   * The value of this constant is 344
   * </p>
   */
  msoElementSeriesAxisGridLinesNone(344),
  /**
   * <p>
   * The value of this constant is 345
   * </p>
   */
  msoElementSeriesAxisGridLinesMinor(345),
  /**
   * <p>
   * The value of this constant is 346
   * </p>
   */
  msoElementSeriesAxisGridLinesMajor(346),
  /**
   * <p>
   * The value of this constant is 347
   * </p>
   */
  msoElementSeriesAxisGridLinesMinorMajor(347),
  /**
   * <p>
   * The value of this constant is 348
   * </p>
   */
  msoElementPrimaryCategoryAxisNone(348),
  /**
   * <p>
   * The value of this constant is 349
   * </p>
   */
  msoElementPrimaryCategoryAxisShow(349),
  /**
   * <p>
   * The value of this constant is 350
   * </p>
   */
  msoElementPrimaryCategoryAxisWithoutLabels(350),
  /**
   * <p>
   * The value of this constant is 351
   * </p>
   */
  msoElementPrimaryCategoryAxisReverse(351),
  /**
   * <p>
   * The value of this constant is 352
   * </p>
   */
  msoElementPrimaryValueAxisNone(352),
  /**
   * <p>
   * The value of this constant is 353
   * </p>
   */
  msoElementPrimaryValueAxisShow(353),
  /**
   * <p>
   * The value of this constant is 354
   * </p>
   */
  msoElementPrimaryValueAxisThousands(354),
  /**
   * <p>
   * The value of this constant is 355
   * </p>
   */
  msoElementPrimaryValueAxisMillions(355),
  /**
   * <p>
   * The value of this constant is 356
   * </p>
   */
  msoElementPrimaryValueAxisBillions(356),
  /**
   * <p>
   * The value of this constant is 357
   * </p>
   */
  msoElementPrimaryValueAxisLogScale(357),
  /**
   * <p>
   * The value of this constant is 358
   * </p>
   */
  msoElementSecondaryCategoryAxisNone(358),
  /**
   * <p>
   * The value of this constant is 359
   * </p>
   */
  msoElementSecondaryCategoryAxisShow(359),
  /**
   * <p>
   * The value of this constant is 360
   * </p>
   */
  msoElementSecondaryCategoryAxisWithoutLabels(360),
  /**
   * <p>
   * The value of this constant is 361
   * </p>
   */
  msoElementSecondaryCategoryAxisReverse(361),
  /**
   * <p>
   * The value of this constant is 362
   * </p>
   */
  msoElementSecondaryValueAxisNone(362),
  /**
   * <p>
   * The value of this constant is 363
   * </p>
   */
  msoElementSecondaryValueAxisShow(363),
  /**
   * <p>
   * The value of this constant is 364
   * </p>
   */
  msoElementSecondaryValueAxisThousands(364),
  /**
   * <p>
   * The value of this constant is 365
   * </p>
   */
  msoElementSecondaryValueAxisMillions(365),
  /**
   * <p>
   * The value of this constant is 366
   * </p>
   */
  msoElementSecondaryValueAxisBillions(366),
  /**
   * <p>
   * The value of this constant is 367
   * </p>
   */
  msoElementSecondaryValueAxisLogScale(367),
  /**
   * <p>
   * The value of this constant is 368
   * </p>
   */
  msoElementSeriesAxisNone(368),
  /**
   * <p>
   * The value of this constant is 369
   * </p>
   */
  msoElementSeriesAxisShow(369),
  /**
   * <p>
   * The value of this constant is 370
   * </p>
   */
  msoElementSeriesAxisWithoutLabeling(370),
  /**
   * <p>
   * The value of this constant is 371
   * </p>
   */
  msoElementSeriesAxisReverse(371),
  /**
   * <p>
   * The value of this constant is 372
   * </p>
   */
  msoElementPrimaryCategoryAxisThousands(372),
  /**
   * <p>
   * The value of this constant is 373
   * </p>
   */
  msoElementPrimaryCategoryAxisMillions(373),
  /**
   * <p>
   * The value of this constant is 374
   * </p>
   */
  msoElementPrimaryCategoryAxisBillions(374),
  /**
   * <p>
   * The value of this constant is 375
   * </p>
   */
  msoElementPrimaryCategoryAxisLogScale(375),
  /**
   * <p>
   * The value of this constant is 376
   * </p>
   */
  msoElementSecondaryCategoryAxisThousands(376),
  /**
   * <p>
   * The value of this constant is 377
   * </p>
   */
  msoElementSecondaryCategoryAxisMillions(377),
  /**
   * <p>
   * The value of this constant is 378
   * </p>
   */
  msoElementSecondaryCategoryAxisBillions(378),
  /**
   * <p>
   * The value of this constant is 379
   * </p>
   */
  msoElementSecondaryCategoryAxisLogScale(379),
  /**
   * <p>
   * The value of this constant is 500
   * </p>
   */
  msoElementDataTableNone(500),
  /**
   * <p>
   * The value of this constant is 501
   * </p>
   */
  msoElementDataTableShow(501),
  /**
   * <p>
   * The value of this constant is 502
   * </p>
   */
  msoElementDataTableWithLegendKeys(502),
  /**
   * <p>
   * The value of this constant is 600
   * </p>
   */
  msoElementTrendlineNone(600),
  /**
   * <p>
   * The value of this constant is 601
   * </p>
   */
  msoElementTrendlineAddLinear(601),
  /**
   * <p>
   * The value of this constant is 602
   * </p>
   */
  msoElementTrendlineAddExponential(602),
  /**
   * <p>
   * The value of this constant is 603
   * </p>
   */
  msoElementTrendlineAddLinearForecast(603),
  /**
   * <p>
   * The value of this constant is 604
   * </p>
   */
  msoElementTrendlineAddTwoPeriodMovingAverage(604),
  /**
   * <p>
   * The value of this constant is 700
   * </p>
   */
  msoElementErrorBarNone(700),
  /**
   * <p>
   * The value of this constant is 701
   * </p>
   */
  msoElementErrorBarStandardError(701),
  /**
   * <p>
   * The value of this constant is 702
   * </p>
   */
  msoElementErrorBarPercentage(702),
  /**
   * <p>
   * The value of this constant is 703
   * </p>
   */
  msoElementErrorBarStandardDeviation(703),
  /**
   * <p>
   * The value of this constant is 800
   * </p>
   */
  msoElementLineNone(800),
  /**
   * <p>
   * The value of this constant is 801
   * </p>
   */
  msoElementLineDropLine(801),
  /**
   * <p>
   * The value of this constant is 802
   * </p>
   */
  msoElementLineHiLoLine(802),
  /**
   * <p>
   * The value of this constant is 803
   * </p>
   */
  msoElementLineSeriesLine(803),
  /**
   * <p>
   * The value of this constant is 804
   * </p>
   */
  msoElementLineDropHiLoLine(804),
  /**
   * <p>
   * The value of this constant is 900
   * </p>
   */
  msoElementUpDownBarsNone(900),
  /**
   * <p>
   * The value of this constant is 901
   * </p>
   */
  msoElementUpDownBarsShow(901),
  /**
   * <p>
   * The value of this constant is 1000
   * </p>
   */
  msoElementPlotAreaNone(1000),
  /**
   * <p>
   * The value of this constant is 1001
   * </p>
   */
  msoElementPlotAreaShow(1001),
  /**
   * <p>
   * The value of this constant is 1100
   * </p>
   */
  msoElementChartWallNone(1100),
  /**
   * <p>
   * The value of this constant is 1101
   * </p>
   */
  msoElementChartWallShow(1101),
  /**
   * <p>
   * The value of this constant is 1200
   * </p>
   */
  msoElementChartFloorNone(1200),
  /**
   * <p>
   * The value of this constant is 1201
   * </p>
   */
  msoElementChartFloorShow(1201),
  ;

  private final int value;
  MsoChartElementType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
