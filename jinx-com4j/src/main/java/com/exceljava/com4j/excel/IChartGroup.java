package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020859-0001-0000-C000-000000000046}")
public interface IChartGroup extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AxisGroup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlAxisGroup
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlAxisGroup getAxisGroup();


  /**
   * <p>
   * Setter method for the COM property "AxisGroup"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlAxisGroup parameter.
   */

  @VTID(11)
  void setAxisGroup(
    com.exceljava.com4j.excel.XlAxisGroup rhs);


  /**
   * <p>
   * Getter method for the COM property "DoughnutHoleSize"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getDoughnutHoleSize();


  /**
   * <p>
   * Setter method for the COM property "DoughnutHoleSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(13)
  void setDoughnutHoleSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "DownBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DownBars
   */

  @VTID(14)
  com.exceljava.com4j.excel.DownBars getDownBars();


  /**
   * <p>
   * Getter method for the COM property "DropLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DropLines
   */

  @VTID(15)
  com.exceljava.com4j.excel.DropLines getDropLines();


  /**
   * <p>
   * Getter method for the COM property "FirstSliceAngle"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getFirstSliceAngle();


  /**
   * <p>
   * Setter method for the COM property "FirstSliceAngle"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(17)
  void setFirstSliceAngle(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "GapWidth"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(18)
  int getGapWidth();


  /**
   * <p>
   * Setter method for the COM property "GapWidth"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(19)
  void setGapWidth(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDropLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getHasDropLines();


  /**
   * <p>
   * Setter method for the COM property "HasDropLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setHasDropLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasHiLoLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getHasHiLoLines();


  /**
   * <p>
   * Setter method for the COM property "HasHiLoLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setHasHiLoLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasRadarAxisLabels"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getHasRadarAxisLabels();


  /**
   * <p>
   * Setter method for the COM property "HasRadarAxisLabels"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setHasRadarAxisLabels(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasSeriesLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getHasSeriesLines();


  /**
   * <p>
   * Setter method for the COM property "HasSeriesLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setHasSeriesLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasUpDownBars"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getHasUpDownBars();


  /**
   * <p>
   * Setter method for the COM property "HasUpDownBars"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setHasUpDownBars(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HiLoLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.HiLoLines
   */

  @VTID(30)
  com.exceljava.com4j.excel.HiLoLines getHiLoLines();


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(31)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Overlap"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(32)
  int getOverlap();


  /**
   * <p>
   * Setter method for the COM property "Overlap"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(33)
  void setOverlap(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "RadarAxisLabels"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TickLabels
   */

  @VTID(34)
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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject seriesCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(35)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject seriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "SeriesLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SeriesLines
   */

  @VTID(36)
  com.exceljava.com4j.excel.SeriesLines getSeriesLines();


  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(37)
  int getSubType();


  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(38)
  void setSubType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(39)
  int getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(40)
  void setType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "UpBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.UpBars
   */

  @VTID(41)
  com.exceljava.com4j.excel.UpBars getUpBars();


  /**
   * <p>
   * Getter method for the COM property "VaryByCategories"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(42)
  boolean getVaryByCategories();


  /**
   * <p>
   * Setter method for the COM property "VaryByCategories"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(43)
  void setVaryByCategories(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SizeRepresents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSizeRepresents
   */

  @VTID(44)
  com.exceljava.com4j.excel.XlSizeRepresents getSizeRepresents();


  /**
   * <p>
   * Setter method for the COM property "SizeRepresents"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSizeRepresents parameter.
   */

  @VTID(45)
  void setSizeRepresents(
    com.exceljava.com4j.excel.XlSizeRepresents rhs);


  /**
   * <p>
   * Getter method for the COM property "BubbleScale"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(46)
  int getBubbleScale();


  /**
   * <p>
   * Setter method for the COM property "BubbleScale"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(47)
  void setBubbleScale(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowNegativeBubbles"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(48)
  boolean getShowNegativeBubbles();


  /**
   * <p>
   * Setter method for the COM property "ShowNegativeBubbles"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(49)
  void setShowNegativeBubbles(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlChartSplitType
   */

  @VTID(50)
  com.exceljava.com4j.excel.XlChartSplitType getSplitType();


  /**
   * <p>
   * Setter method for the COM property "SplitType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartSplitType parameter.
   */

  @VTID(51)
  void setSplitType(
    com.exceljava.com4j.excel.XlChartSplitType rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitValue"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSplitValue();


  /**
   * <p>
   * Setter method for the COM property "SplitValue"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  void setSplitValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SecondPlotSize"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(54)
  int getSecondPlotSize();


  /**
   * <p>
   * Setter method for the COM property "SecondPlotSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(55)
  void setSecondPlotSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Has3DShading"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(56)
  boolean getHas3DShading();


  /**
   * <p>
   * Setter method for the COM property "Has3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(57)
  void setHas3DShading(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullCategoryCollection(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject fullCategoryCollection();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullCategoryCollection(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject fullCategoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(58)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject fullCategoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * categoryCollection(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject categoryCollection();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * categoryCollection(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject categoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(59)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject categoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "BinsType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlBinsType
   */

  @VTID(60)
  com.exceljava.com4j.excel.XlBinsType getBinsType();


  /**
   * <p>
   * Setter method for the COM property "BinsType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlBinsType parameter.
   */

  @VTID(61)
  void setBinsType(
    com.exceljava.com4j.excel.XlBinsType rhs);


  /**
   * <p>
   * Getter method for the COM property "BinWidthValue"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(62)
  double getBinWidthValue();


  /**
   * <p>
   * Setter method for the COM property "BinWidthValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(63)
  void setBinWidthValue(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsCountValue"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(64)
  int getBinsCountValue();


  /**
   * <p>
   * Setter method for the COM property "BinsCountValue"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(65)
  void setBinsCountValue(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowEnabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(66)
  boolean getBinsOverflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowEnabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(67)
  void setBinsOverflowEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowValue"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(68)
  double getBinsOverflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(69)
  void setBinsOverflowValue(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowEnabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(70)
  boolean getBinsUnderflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowEnabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(71)
  void setBinsUnderflowEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowValue"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(72)
  double getBinsUnderflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowValue"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(73)
  void setBinsUnderflowValue(
    double rhs);


  // Properties:
}
