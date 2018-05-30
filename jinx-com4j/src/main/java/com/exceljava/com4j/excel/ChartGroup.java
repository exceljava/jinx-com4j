package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ChartGroup extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AxisGroup"
   * </p>
   */

  @DISPID(47)
  @PropGet
  com.exceljava.com4j.excel.XlAxisGroup getAxisGroup();


  /**
   * <p>
   * Setter method for the COM property "AxisGroup"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlAxisGroup parameter.
   */

  @DISPID(47)
  @PropPut
  void setAxisGroup(
    com.exceljava.com4j.excel.XlAxisGroup rhs);


  /**
   * <p>
   * Getter method for the COM property "DoughnutHoleSize"
   * </p>
   */

  @DISPID(1126)
  @PropGet
  int getDoughnutHoleSize();


  /**
   * <p>
   * Setter method for the COM property "DoughnutHoleSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1126)
  @PropPut
  void setDoughnutHoleSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "DownBars"
   * </p>
   */

  @DISPID(141)
  @PropGet
  com.exceljava.com4j.excel.DownBars getDownBars();


  /**
   * <p>
   * Getter method for the COM property "DropLines"
   * </p>
   */

  @DISPID(142)
  @PropGet
  com.exceljava.com4j.excel.DropLines getDropLines();


  /**
   * <p>
   * Getter method for the COM property "FirstSliceAngle"
   * </p>
   */

  @DISPID(63)
  @PropGet
  int getFirstSliceAngle();


  /**
   * <p>
   * Setter method for the COM property "FirstSliceAngle"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(63)
  @PropPut
  void setFirstSliceAngle(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "GapWidth"
   * </p>
   */

  @DISPID(51)
  @PropGet
  int getGapWidth();


  /**
   * <p>
   * Setter method for the COM property "GapWidth"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(51)
  @PropPut
  void setGapWidth(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDropLines"
   * </p>
   */

  @DISPID(61)
  @PropGet
  boolean getHasDropLines();


  /**
   * <p>
   * Setter method for the COM property "HasDropLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(61)
  @PropPut
  void setHasDropLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasHiLoLines"
   * </p>
   */

  @DISPID(62)
  @PropGet
  boolean getHasHiLoLines();


  /**
   * <p>
   * Setter method for the COM property "HasHiLoLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(62)
  @PropPut
  void setHasHiLoLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasRadarAxisLabels"
   * </p>
   */

  @DISPID(64)
  @PropGet
  boolean getHasRadarAxisLabels();


  /**
   * <p>
   * Setter method for the COM property "HasRadarAxisLabels"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(64)
  @PropPut
  void setHasRadarAxisLabels(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasSeriesLines"
   * </p>
   */

  @DISPID(65)
  @PropGet
  boolean getHasSeriesLines();


  /**
   * <p>
   * Setter method for the COM property "HasSeriesLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(65)
  @PropPut
  void setHasSeriesLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasUpDownBars"
   * </p>
   */

  @DISPID(66)
  @PropGet
  boolean getHasUpDownBars();


  /**
   * <p>
   * Setter method for the COM property "HasUpDownBars"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(66)
  @PropPut
  void setHasUpDownBars(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HiLoLines"
   * </p>
   */

  @DISPID(143)
  @PropGet
  com.exceljava.com4j.excel.HiLoLines getHiLoLines();


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   */

  @DISPID(486)
  @PropGet
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Overlap"
   * </p>
   */

  @DISPID(56)
  @PropGet
  int getOverlap();


  /**
   * <p>
   * Setter method for the COM property "Overlap"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(56)
  @PropPut
  void setOverlap(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "RadarAxisLabels"
   * </p>
   */

  @DISPID(144)
  @PropGet
  com.exceljava.com4j.excel.TickLabels getRadarAxisLabels();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * seriesCollection(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(68)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject seriesCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(68)
  com4j.Com4jObject seriesCollection(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "SeriesLines"
   * </p>
   */

  @DISPID(145)
  @PropGet
  com.exceljava.com4j.excel.SeriesLines getSeriesLines();


  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   */

  @DISPID(109)
  @PropGet
  int getSubType();


  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(109)
  @PropPut
  void setSubType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  int getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(108)
  @PropPut
  void setType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "UpBars"
   * </p>
   */

  @DISPID(140)
  @PropGet
  com.exceljava.com4j.excel.UpBars getUpBars();


  /**
   * <p>
   * Getter method for the COM property "VaryByCategories"
   * </p>
   */

  @DISPID(60)
  @PropGet
  boolean getVaryByCategories();


  /**
   * <p>
   * Setter method for the COM property "VaryByCategories"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(60)
  @PropPut
  void setVaryByCategories(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SizeRepresents"
   * </p>
   */

  @DISPID(1652)
  @PropGet
  com.exceljava.com4j.excel.XlSizeRepresents getSizeRepresents();


  /**
   * <p>
   * Setter method for the COM property "SizeRepresents"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSizeRepresents parameter.
   */

  @DISPID(1652)
  @PropPut
  void setSizeRepresents(
    com.exceljava.com4j.excel.XlSizeRepresents rhs);


  /**
   * <p>
   * Getter method for the COM property "BubbleScale"
   * </p>
   */

  @DISPID(1653)
  @PropGet
  int getBubbleScale();


  /**
   * <p>
   * Setter method for the COM property "BubbleScale"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1653)
  @PropPut
  void setBubbleScale(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowNegativeBubbles"
   * </p>
   */

  @DISPID(1654)
  @PropGet
  boolean getShowNegativeBubbles();


  /**
   * <p>
   * Setter method for the COM property "ShowNegativeBubbles"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1654)
  @PropPut
  void setShowNegativeBubbles(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitType"
   * </p>
   */

  @DISPID(1655)
  @PropGet
  com.exceljava.com4j.excel.XlChartSplitType getSplitType();


  /**
   * <p>
   * Setter method for the COM property "SplitType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartSplitType parameter.
   */

  @DISPID(1655)
  @PropPut
  void setSplitType(
    com.exceljava.com4j.excel.XlChartSplitType rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitValue"
   * </p>
   */

  @DISPID(1656)
  @PropGet
  java.lang.Object getSplitValue();


  /**
   * <p>
   * Setter method for the COM property "SplitValue"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1656)
  @PropPut
  void setSplitValue(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SecondPlotSize"
   * </p>
   */

  @DISPID(1657)
  @PropGet
  int getSecondPlotSize();


  /**
   * <p>
   * Setter method for the COM property "SecondPlotSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1657)
  @PropPut
  void setSecondPlotSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Has3DShading"
   * </p>
   */

  @DISPID(1658)
  @PropGet
  boolean getHas3DShading();


  /**
   * <p>
   * Setter method for the COM property "Has3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1658)
  @PropPut
  void setHas3DShading(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullCategoryCollection(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3081)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject fullCategoryCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3081)
  com4j.Com4jObject fullCategoryCollection(
    @Optional java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * categoryCollection(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(3082)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject categoryCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3082)
  com4j.Com4jObject categoryCollection(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "BinsType"
   * </p>
   */

  @DISPID(3196)
  @PropGet
  com.exceljava.com4j.excel.XlBinsType getBinsType();


  /**
   * <p>
   * Setter method for the COM property "BinsType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlBinsType parameter.
   */

  @DISPID(3196)
  @PropPut
  void setBinsType(
    com.exceljava.com4j.excel.XlBinsType rhs);


  /**
   * <p>
   * Getter method for the COM property "BinWidthValue"
   * </p>
   */

  @DISPID(3197)
  @PropGet
  double getBinWidthValue();


  /**
   * <p>
   * Setter method for the COM property "BinWidthValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(3197)
  @PropPut
  void setBinWidthValue(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsCountValue"
   * </p>
   */

  @DISPID(3198)
  @PropGet
  int getBinsCountValue();


  /**
   * <p>
   * Setter method for the COM property "BinsCountValue"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(3198)
  @PropPut
  void setBinsCountValue(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowEnabled"
   * </p>
   */

  @DISPID(3199)
  @PropGet
  boolean getBinsOverflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowEnabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3199)
  @PropPut
  void setBinsOverflowEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowValue"
   * </p>
   */

  @DISPID(3200)
  @PropGet
  double getBinsOverflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(3200)
  @PropPut
  void setBinsOverflowValue(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowEnabled"
   * </p>
   */

  @DISPID(3201)
  @PropGet
  boolean getBinsUnderflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowEnabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3201)
  @PropPut
  void setBinsUnderflowEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowValue"
   * </p>
   */

  @DISPID(3202)
  @PropGet
  double getBinsUnderflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(3202)
  @PropPut
  void setBinsUnderflowValue(
    double rhs);


  // Properties:
}
