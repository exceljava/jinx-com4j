package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1727-0000-0000-C000-000000000046}")
public interface IMsoChartGroup extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Setter method for the COM property "AxisGroup"
   * </p>
   * @param piGroup Mandatory int parameter.
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  void setAxisGroup(
    int piGroup);


  /**
   * <p>
   * Getter method for the COM property "AxisGroup"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(8)
  int getAxisGroup();


  /**
   * <p>
   * Setter method for the COM property "DoughnutHoleSize"
   * </p>
   * @param pDoughnutHoleSize Mandatory int parameter.
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  void setDoughnutHoleSize(
    int pDoughnutHoleSize);


  /**
   * <p>
   * Getter method for the COM property "DoughnutHoleSize"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(10)
  int getDoughnutHoleSize();


  /**
   * <p>
   * Getter method for the COM property "DownBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDownBars
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.IMsoDownBars getDownBars();


  /**
   * <p>
   * Getter method for the COM property "DropLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDropLines
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.IMsoDropLines getDropLines();


  /**
   * <p>
   * Setter method for the COM property "FirstSliceAngle"
   * </p>
   * @param pFirstSliceAngle Mandatory int parameter.
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  void setFirstSliceAngle(
    int pFirstSliceAngle);


  /**
   * <p>
   * Getter method for the COM property "FirstSliceAngle"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(14)
  int getFirstSliceAngle();


  /**
   * <p>
   * Setter method for the COM property "GapWidth"
   * </p>
   * @param pGapWidth Mandatory int parameter.
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  void setGapWidth(
    int pGapWidth);


  /**
   * <p>
   * Getter method for the COM property "GapWidth"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(16)
  int getGapWidth();


  /**
   * <p>
   * Setter method for the COM property "HasDropLines"
   * </p>
   * @param pfHasDropLines Mandatory boolean parameter.
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  void setHasDropLines(
    boolean pfHasDropLines);


  /**
   * <p>
   * Getter method for the COM property "HasDropLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(18)
  boolean getHasDropLines();


  /**
   * <p>
   * Setter method for the COM property "HasHiLoLines"
   * </p>
   * @param pfHasHiLoLines Mandatory boolean parameter.
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(19)
  void setHasHiLoLines(
    boolean pfHasHiLoLines);


  /**
   * <p>
   * Getter method for the COM property "HasHiLoLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(20)
  boolean getHasHiLoLines();


  /**
   * <p>
   * Setter method for the COM property "HasRadarAxisLabels"
   * </p>
   * @param pfHasRadarAxisLabels Mandatory boolean parameter.
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(21)
  void setHasRadarAxisLabels(
    boolean pfHasRadarAxisLabels);


  /**
   * <p>
   * Getter method for the COM property "HasRadarAxisLabels"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(22)
  boolean getHasRadarAxisLabels();


  /**
   * <p>
   * Setter method for the COM property "HasSeriesLines"
   * </p>
   * @param pfHasSeriesLines Mandatory boolean parameter.
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(23)
  void setHasSeriesLines(
    boolean pfHasSeriesLines);


  /**
   * <p>
   * Getter method for the COM property "HasSeriesLines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(24)
  boolean getHasSeriesLines();


  /**
   * <p>
   * Setter method for the COM property "HasUpDownBars"
   * </p>
   * @param pfHasUpDownBars Mandatory boolean parameter.
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(25)
  void setHasUpDownBars(
    boolean pfHasUpDownBars);


  /**
   * <p>
   * Getter method for the COM property "HasUpDownBars"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(26)
  boolean getHasUpDownBars();


  /**
   * <p>
   * Getter method for the COM property "HiLoLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoHiLoLines
   */

  @DISPID(1610743828) //= 0x60020014. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.IMsoHiLoLines getHiLoLines();


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743829) //= 0x60020015. The runtime will prefer the VTID if present
  @VTID(28)
  int getIndex();


  /**
   * <p>
   * Setter method for the COM property "Overlap"
   * </p>
   * @param pOverlap Mandatory int parameter.
   */

  @DISPID(1610743830) //= 0x60020016. The runtime will prefer the VTID if present
  @VTID(29)
  void setOverlap(
    int pOverlap);


  /**
   * <p>
   * Getter method for the COM property "Overlap"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743830) //= 0x60020016. The runtime will prefer the VTID if present
  @VTID(30)
  int getOverlap();


  /**
   * <p>
   * Getter method for the COM property "RadarAxisLabels"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743832) //= 0x60020018. The runtime will prefer the VTID if present
  @VTID(31)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getRadarAxisLabels();


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

  @DISPID(1610743833) //= 0x60020019. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject seriesCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743833) //= 0x60020019. The runtime will prefer the VTID if present
  @VTID(32)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject seriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "SeriesLines"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeriesLines
   */

  @DISPID(1610743834) //= 0x6002001a. The runtime will prefer the VTID if present
  @VTID(33)
  com.exceljava.com4j.office.IMsoSeriesLines getSeriesLines();


  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * @param pSubType Mandatory int parameter.
   */

  @DISPID(1610743835) //= 0x6002001b. The runtime will prefer the VTID if present
  @VTID(34)
  void setSubType(
    int pSubType);


  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743835) //= 0x6002001b. The runtime will prefer the VTID if present
  @VTID(35)
  int getSubType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param ptype Mandatory int parameter.
   */

  @DISPID(1610743837) //= 0x6002001d. The runtime will prefer the VTID if present
  @VTID(36)
  void setType(
    int ptype);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743837) //= 0x6002001d. The runtime will prefer the VTID if present
  @VTID(37)
  int getType();


  /**
   * <p>
   * Getter method for the COM property "UpBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoUpBars
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.IMsoUpBars getUpBars();


  /**
   * <p>
   * Setter method for the COM property "VaryByCategories"
   * </p>
   * @param pfVaryByCategories Mandatory boolean parameter.
   */

  @DISPID(1610743840) //= 0x60020020. The runtime will prefer the VTID if present
  @VTID(39)
  void setVaryByCategories(
    boolean pfVaryByCategories);


  /**
   * <p>
   * Getter method for the COM property "VaryByCategories"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743840) //= 0x60020020. The runtime will prefer the VTID if present
  @VTID(40)
  boolean getVaryByCategories();


  /**
   * <p>
   * Getter method for the COM property "SizeRepresents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlSizeRepresents
   */

  @DISPID(1610743842) //= 0x60020022. The runtime will prefer the VTID if present
  @VTID(41)
  com.exceljava.com4j.office.XlSizeRepresents getSizeRepresents();


  /**
   * <p>
   * Setter method for the COM property "SizeRepresents"
   * </p>
   * @param pXlSizeRepresents Mandatory com.exceljava.com4j.office.XlSizeRepresents parameter.
   */

  @DISPID(1610743842) //= 0x60020022. The runtime will prefer the VTID if present
  @VTID(42)
  void setSizeRepresents(
    com.exceljava.com4j.office.XlSizeRepresents pXlSizeRepresents);


  /**
   * <p>
   * Setter method for the COM property "BubbleScale"
   * </p>
   * @param pbubblescale Mandatory int parameter.
   */

  @DISPID(1610743844) //= 0x60020024. The runtime will prefer the VTID if present
  @VTID(43)
  void setBubbleScale(
    int pbubblescale);


  /**
   * <p>
   * Getter method for the COM property "BubbleScale"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743844) //= 0x60020024. The runtime will prefer the VTID if present
  @VTID(44)
  int getBubbleScale();


  /**
   * <p>
   * Setter method for the COM property "ShowNegativeBubbles"
   * </p>
   * @param pfShowNegativeBubbles Mandatory boolean parameter.
   */

  @DISPID(1610743846) //= 0x60020026. The runtime will prefer the VTID if present
  @VTID(45)
  void setShowNegativeBubbles(
    boolean pfShowNegativeBubbles);


  /**
   * <p>
   * Getter method for the COM property "ShowNegativeBubbles"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743846) //= 0x60020026. The runtime will prefer the VTID if present
  @VTID(46)
  boolean getShowNegativeBubbles();


  /**
   * <p>
   * Setter method for the COM property "SplitType"
   * </p>
   * @param pChartSplitType Mandatory com.exceljava.com4j.office.XlChartSplitType parameter.
   */

  @DISPID(1610743848) //= 0x60020028. The runtime will prefer the VTID if present
  @VTID(47)
  void setSplitType(
    com.exceljava.com4j.office.XlChartSplitType pChartSplitType);


  /**
   * <p>
   * Getter method for the COM property "SplitType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlChartSplitType
   */

  @DISPID(1610743848) //= 0x60020028. The runtime will prefer the VTID if present
  @VTID(48)
  com.exceljava.com4j.office.XlChartSplitType getSplitType();


  /**
   * <p>
   * Getter method for the COM property "SplitValue"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743850) //= 0x6002002a. The runtime will prefer the VTID if present
  @VTID(49)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSplitValue();


  /**
   * <p>
   * Setter method for the COM property "SplitValue"
   * </p>
   * @param pSplitValue Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743850) //= 0x6002002a. The runtime will prefer the VTID if present
  @VTID(50)
  void setSplitValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pSplitValue);


  /**
   * <p>
   * Getter method for the COM property "SecondPlotSize"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743852) //= 0x6002002c. The runtime will prefer the VTID if present
  @VTID(51)
  int getSecondPlotSize();


  /**
   * <p>
   * Setter method for the COM property "SecondPlotSize"
   * </p>
   * @param pSecondPlotSize Mandatory int parameter.
   */

  @DISPID(1610743852) //= 0x6002002c. The runtime will prefer the VTID if present
  @VTID(52)
  void setSecondPlotSize(
    int pSecondPlotSize);


  /**
   * <p>
   * Getter method for the COM property "Has3DShading"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743854) //= 0x6002002e. The runtime will prefer the VTID if present
  @VTID(53)
  boolean getHas3DShading();


  /**
   * <p>
   * Setter method for the COM property "Has3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1610743854) //= 0x6002002e. The runtime will prefer the VTID if present
  @VTID(54)
  void setHas3DShading(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(55)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(56)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(57)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject categoryCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(58)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject categoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(152) //= 0x98. The runtime will prefer the VTID if present
  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject fullCategoryCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(152) //= 0x98. The runtime will prefer the VTID if present
  @VTID(59)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject fullCategoryCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "BinsType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlBinsType
   */

  @DISPID(2891) //= 0xb4b. The runtime will prefer the VTID if present
  @VTID(60)
  com.exceljava.com4j.office.XlBinsType getBinsType();


  /**
   * <p>
   * Setter method for the COM property "BinsType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.XlBinsType parameter.
   */

  @DISPID(2891) //= 0xb4b. The runtime will prefer the VTID if present
  @VTID(61)
  void setBinsType(
    com.exceljava.com4j.office.XlBinsType rhs);


  /**
   * <p>
   * Getter method for the COM property "BinWidthValue"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(2892) //= 0xb4c. The runtime will prefer the VTID if present
  @VTID(62)
  double getBinWidthValue();


  /**
   * <p>
   * Setter method for the COM property "BinWidthValue"
   * </p>
   * @param pval Mandatory double parameter.
   */

  @DISPID(2892) //= 0xb4c. The runtime will prefer the VTID if present
  @VTID(63)
  void setBinWidthValue(
    double pval);


  /**
   * <p>
   * Getter method for the COM property "BinsCountValue"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2893) //= 0xb4d. The runtime will prefer the VTID if present
  @VTID(64)
  int getBinsCountValue();


  /**
   * <p>
   * Setter method for the COM property "BinsCountValue"
   * </p>
   * @param pval Mandatory int parameter.
   */

  @DISPID(2893) //= 0xb4d. The runtime will prefer the VTID if present
  @VTID(65)
  void setBinsCountValue(
    int pval);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowEnabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2894) //= 0xb4e. The runtime will prefer the VTID if present
  @VTID(66)
  boolean getBinsOverflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowEnabled"
   * </p>
   * @param pval Mandatory boolean parameter.
   */

  @DISPID(2894) //= 0xb4e. The runtime will prefer the VTID if present
  @VTID(67)
  void setBinsOverflowEnabled(
    boolean pval);


  /**
   * <p>
   * Getter method for the COM property "BinsOverflowValue"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(2895) //= 0xb4f. The runtime will prefer the VTID if present
  @VTID(68)
  double getBinsOverflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsOverflowValue"
   * </p>
   * @param pval Mandatory double parameter.
   */

  @DISPID(2895) //= 0xb4f. The runtime will prefer the VTID if present
  @VTID(69)
  void setBinsOverflowValue(
    double pval);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowEnabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2896) //= 0xb50. The runtime will prefer the VTID if present
  @VTID(70)
  boolean getBinsUnderflowEnabled();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowEnabled"
   * </p>
   * @param pval Mandatory boolean parameter.
   */

  @DISPID(2896) //= 0xb50. The runtime will prefer the VTID if present
  @VTID(71)
  void setBinsUnderflowEnabled(
    boolean pval);


  /**
   * <p>
   * Getter method for the COM property "BinsUnderflowValue"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(2897) //= 0xb51. The runtime will prefer the VTID if present
  @VTID(72)
  double getBinsUnderflowValue();


  /**
   * <p>
   * Setter method for the COM property "BinsUnderflowValue"
   * </p>
   * @param pval Mandatory double parameter.
   */

  @DISPID(2897) //= 0xb51. The runtime will prefer the VTID if present
  @VTID(73)
  void setBinsUnderflowValue(
    double pval);


  // Properties:
}
