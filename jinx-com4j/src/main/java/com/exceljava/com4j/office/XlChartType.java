package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlChartType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  xlColumnClustered(51),
  /**
   * <p>
   * The value of this constant is 52
   * </p>
   */
  xlColumnStacked(52),
  /**
   * <p>
   * The value of this constant is 53
   * </p>
   */
  xlColumnStacked100(53),
  /**
   * <p>
   * The value of this constant is 54
   * </p>
   */
  xl3DColumnClustered(54),
  /**
   * <p>
   * The value of this constant is 55
   * </p>
   */
  xl3DColumnStacked(55),
  /**
   * <p>
   * The value of this constant is 56
   * </p>
   */
  xl3DColumnStacked100(56),
  /**
   * <p>
   * The value of this constant is 57
   * </p>
   */
  xlBarClustered(57),
  /**
   * <p>
   * The value of this constant is 58
   * </p>
   */
  xlBarStacked(58),
  /**
   * <p>
   * The value of this constant is 59
   * </p>
   */
  xlBarStacked100(59),
  /**
   * <p>
   * The value of this constant is 60
   * </p>
   */
  xl3DBarClustered(60),
  /**
   * <p>
   * The value of this constant is 61
   * </p>
   */
  xl3DBarStacked(61),
  /**
   * <p>
   * The value of this constant is 62
   * </p>
   */
  xl3DBarStacked100(62),
  /**
   * <p>
   * The value of this constant is 63
   * </p>
   */
  xlLineStacked(63),
  /**
   * <p>
   * The value of this constant is 64
   * </p>
   */
  xlLineStacked100(64),
  /**
   * <p>
   * The value of this constant is 65
   * </p>
   */
  xlLineMarkers(65),
  /**
   * <p>
   * The value of this constant is 66
   * </p>
   */
  xlLineMarkersStacked(66),
  /**
   * <p>
   * The value of this constant is 67
   * </p>
   */
  xlLineMarkersStacked100(67),
  /**
   * <p>
   * The value of this constant is 68
   * </p>
   */
  xlPieOfPie(68),
  /**
   * <p>
   * The value of this constant is 69
   * </p>
   */
  xlPieExploded(69),
  /**
   * <p>
   * The value of this constant is 70
   * </p>
   */
  xl3DPieExploded(70),
  /**
   * <p>
   * The value of this constant is 71
   * </p>
   */
  xlBarOfPie(71),
  /**
   * <p>
   * The value of this constant is 72
   * </p>
   */
  xlXYScatterSmooth(72),
  /**
   * <p>
   * The value of this constant is 73
   * </p>
   */
  xlXYScatterSmoothNoMarkers(73),
  /**
   * <p>
   * The value of this constant is 74
   * </p>
   */
  xlXYScatterLines(74),
  /**
   * <p>
   * The value of this constant is 75
   * </p>
   */
  xlXYScatterLinesNoMarkers(75),
  /**
   * <p>
   * The value of this constant is 76
   * </p>
   */
  xlAreaStacked(76),
  /**
   * <p>
   * The value of this constant is 77
   * </p>
   */
  xlAreaStacked100(77),
  /**
   * <p>
   * The value of this constant is 78
   * </p>
   */
  xl3DAreaStacked(78),
  /**
   * <p>
   * The value of this constant is 79
   * </p>
   */
  xl3DAreaStacked100(79),
  /**
   * <p>
   * The value of this constant is 80
   * </p>
   */
  xlDoughnutExploded(80),
  /**
   * <p>
   * The value of this constant is 81
   * </p>
   */
  xlRadarMarkers(81),
  /**
   * <p>
   * The value of this constant is 82
   * </p>
   */
  xlRadarFilled(82),
  /**
   * <p>
   * The value of this constant is 83
   * </p>
   */
  xlSurface(83),
  /**
   * <p>
   * The value of this constant is 84
   * </p>
   */
  xlSurfaceWireframe(84),
  /**
   * <p>
   * The value of this constant is 85
   * </p>
   */
  xlSurfaceTopView(85),
  /**
   * <p>
   * The value of this constant is 86
   * </p>
   */
  xlSurfaceTopViewWireframe(86),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlBubble(15),
  /**
   * <p>
   * The value of this constant is 87
   * </p>
   */
  xlBubble3DEffect(87),
  /**
   * <p>
   * The value of this constant is 88
   * </p>
   */
  xlStockHLC(88),
  /**
   * <p>
   * The value of this constant is 89
   * </p>
   */
  xlStockOHLC(89),
  /**
   * <p>
   * The value of this constant is 90
   * </p>
   */
  xlStockVHLC(90),
  /**
   * <p>
   * The value of this constant is 91
   * </p>
   */
  xlStockVOHLC(91),
  /**
   * <p>
   * The value of this constant is 92
   * </p>
   */
  xlCylinderColClustered(92),
  /**
   * <p>
   * The value of this constant is 93
   * </p>
   */
  xlCylinderColStacked(93),
  /**
   * <p>
   * The value of this constant is 94
   * </p>
   */
  xlCylinderColStacked100(94),
  /**
   * <p>
   * The value of this constant is 95
   * </p>
   */
  xlCylinderBarClustered(95),
  /**
   * <p>
   * The value of this constant is 96
   * </p>
   */
  xlCylinderBarStacked(96),
  /**
   * <p>
   * The value of this constant is 97
   * </p>
   */
  xlCylinderBarStacked100(97),
  /**
   * <p>
   * The value of this constant is 98
   * </p>
   */
  xlCylinderCol(98),
  /**
   * <p>
   * The value of this constant is 99
   * </p>
   */
  xlConeColClustered(99),
  /**
   * <p>
   * The value of this constant is 100
   * </p>
   */
  xlConeColStacked(100),
  /**
   * <p>
   * The value of this constant is 101
   * </p>
   */
  xlConeColStacked100(101),
  /**
   * <p>
   * The value of this constant is 102
   * </p>
   */
  xlConeBarClustered(102),
  /**
   * <p>
   * The value of this constant is 103
   * </p>
   */
  xlConeBarStacked(103),
  /**
   * <p>
   * The value of this constant is 104
   * </p>
   */
  xlConeBarStacked100(104),
  /**
   * <p>
   * The value of this constant is 105
   * </p>
   */
  xlConeCol(105),
  /**
   * <p>
   * The value of this constant is 106
   * </p>
   */
  xlPyramidColClustered(106),
  /**
   * <p>
   * The value of this constant is 107
   * </p>
   */
  xlPyramidColStacked(107),
  /**
   * <p>
   * The value of this constant is 108
   * </p>
   */
  xlPyramidColStacked100(108),
  /**
   * <p>
   * The value of this constant is 109
   * </p>
   */
  xlPyramidBarClustered(109),
  /**
   * <p>
   * The value of this constant is 110
   * </p>
   */
  xlPyramidBarStacked(110),
  /**
   * <p>
   * The value of this constant is 111
   * </p>
   */
  xlPyramidBarStacked100(111),
  /**
   * <p>
   * The value of this constant is 112
   * </p>
   */
  xlPyramidCol(112),
  /**
   * <p>
   * The value of this constant is -4100
   * </p>
   */
  xl3DColumn(-4100),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlLine(4),
  /**
   * <p>
   * The value of this constant is -4101
   * </p>
   */
  xl3DLine(-4101),
  /**
   * <p>
   * The value of this constant is -4102
   * </p>
   */
  xl3DPie(-4102),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlPie(5),
  /**
   * <p>
   * The value of this constant is -4169
   * </p>
   */
  xlXYScatter(-4169),
  /**
   * <p>
   * The value of this constant is -4098
   * </p>
   */
  xl3DArea(-4098),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlArea(1),
  /**
   * <p>
   * The value of this constant is -4120
   * </p>
   */
  xlDoughnut(-4120),
  /**
   * <p>
   * The value of this constant is -4151
   * </p>
   */
  xlRadar(-4151),
  /**
   * <p>
   * The value of this constant is -4152
   * </p>
   */
  xlCombo(-4152),
  /**
   * <p>
   * The value of this constant is 113
   * </p>
   */
  xlComboColumnClusteredLine(113),
  /**
   * <p>
   * The value of this constant is 114
   * </p>
   */
  xlComboColumnClusteredLineSecondaryAxis(114),
  /**
   * <p>
   * The value of this constant is 115
   * </p>
   */
  xlComboAreaStackedColumnClustered(115),
  /**
   * <p>
   * The value of this constant is 116
   * </p>
   */
  xlOtherCombinations(116),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  xlSuggestedChart(-2),
  /**
   * <p>
   * The value of this constant is 117
   * </p>
   */
  xlTreemap(117),
  /**
   * <p>
   * The value of this constant is 118
   * </p>
   */
  xlHistogram(118),
  /**
   * <p>
   * The value of this constant is 119
   * </p>
   */
  xlWaterfall(119),
  /**
   * <p>
   * The value of this constant is 120
   * </p>
   */
  xlSunburst(120),
  /**
   * <p>
   * The value of this constant is 121
   * </p>
   */
  xlBoxwhisker(121),
  /**
   * <p>
   * The value of this constant is 122
   * </p>
   */
  xlPareto(122),
  /**
   * <p>
   * The value of this constant is 123
   * </p>
   */
  xlFunnel(123),
  /**
   * <p>
   * The value of this constant is 124
   * </p>
   */
  xlColumnClusteredEx(124),
  /**
   * <p>
   * The value of this constant is 125
   * </p>
   */
  xlColumnStackedEx(125),
  /**
   * <p>
   * The value of this constant is 126
   * </p>
   */
  xlColumnStacked100Ex(126),
  /**
   * <p>
   * The value of this constant is 127
   * </p>
   */
  xlLineEx(127),
  /**
   * <p>
   * The value of this constant is 128
   * </p>
   */
  xlLineStackedEx(128),
  /**
   * <p>
   * The value of this constant is 129
   * </p>
   */
  xlLineStacked100Ex(129),
  /**
   * <p>
   * The value of this constant is 130
   * </p>
   */
  xlPieEx(130),
  /**
   * <p>
   * The value of this constant is 131
   * </p>
   */
  xlDoughnutEx(131),
  /**
   * <p>
   * The value of this constant is 132
   * </p>
   */
  xlBarClusteredEx(132),
  /**
   * <p>
   * The value of this constant is 133
   * </p>
   */
  xlBarStackedEx(133),
  /**
   * <p>
   * The value of this constant is 134
   * </p>
   */
  xlBarStacked100Ex(134),
  /**
   * <p>
   * The value of this constant is 135
   * </p>
   */
  xlAreaEx(135),
  /**
   * <p>
   * The value of this constant is 136
   * </p>
   */
  xlAreaStackedEx(136),
  /**
   * <p>
   * The value of this constant is 137
   * </p>
   */
  xlAreaStacked100Ex(137),
  /**
   * <p>
   * The value of this constant is 138
   * </p>
   */
  xlXYScatterEx(138),
  /**
   * <p>
   * The value of this constant is 139
   * </p>
   */
  xlBubbleEx(139),
  /**
   * <p>
   * The value of this constant is 140
   * </p>
   */
  xlRegionMap(140),
  ;

  private final int value;
  XlChartType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
