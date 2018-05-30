package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Series extends Com4jObject {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(151)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _ApplyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(151)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, legendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, legendKey, autoText, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText);

  /**
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines);


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
   * Getter method for the COM property "Border"
   * </p>
   */

  @DISPID(128)
  @PropGet
  com.exceljava.com4j.excel.Border getBorder();


  /**
   */

  @DISPID(112)
  java.lang.Object clearFormats();


  /**
   */

  @DISPID(551)
  java.lang.Object copy();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dataLabels(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(157)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject dataLabels();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(157)
  com4j.Com4jObject dataLabels(
    @Optional java.lang.Object index);


  /**
   */

  @DISPID(117)
  java.lang.Object delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter amount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter minusValues is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * errorBar(direction, include, type, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param direction Mandatory com.exceljava.com4j.excel.XlErrorBarDirection parameter.
   * @param include Mandatory com.exceljava.com4j.excel.XlErrorBarInclude parameter.
   * @param type Mandatory com.exceljava.com4j.excel.XlErrorBarType parameter.
   */

  @DISPID(152)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object errorBar(
    com.exceljava.com4j.excel.XlErrorBarDirection direction,
    com.exceljava.com4j.excel.XlErrorBarInclude include,
    com.exceljava.com4j.excel.XlErrorBarType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter minusValues is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * errorBar(direction, include, type, amount, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param direction Mandatory com.exceljava.com4j.excel.XlErrorBarDirection parameter.
   * @param include Mandatory com.exceljava.com4j.excel.XlErrorBarInclude parameter.
   * @param type Mandatory com.exceljava.com4j.excel.XlErrorBarType parameter.
   * @param amount Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(152)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object errorBar(
    com.exceljava.com4j.excel.XlErrorBarDirection direction,
    com.exceljava.com4j.excel.XlErrorBarInclude include,
    com.exceljava.com4j.excel.XlErrorBarType type,
    @Optional java.lang.Object amount);

  /**
   * @param direction Mandatory com.exceljava.com4j.excel.XlErrorBarDirection parameter.
   * @param include Mandatory com.exceljava.com4j.excel.XlErrorBarInclude parameter.
   * @param type Mandatory com.exceljava.com4j.excel.XlErrorBarType parameter.
   * @param amount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param minusValues Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(152)
  java.lang.Object errorBar(
    com.exceljava.com4j.excel.XlErrorBarDirection direction,
    com.exceljava.com4j.excel.XlErrorBarInclude include,
    com.exceljava.com4j.excel.XlErrorBarType type,
    @Optional java.lang.Object amount,
    @Optional java.lang.Object minusValues);


  /**
   * <p>
   * Getter method for the COM property "ErrorBars"
   * </p>
   */

  @DISPID(159)
  @PropGet
  com.exceljava.com4j.excel.ErrorBars getErrorBars();


  /**
   * <p>
   * Getter method for the COM property "Explosion"
   * </p>
   */

  @DISPID(182)
  @PropGet
  int getExplosion();


  /**
   * <p>
   * Setter method for the COM property "Explosion"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(182)
  @PropPut
  void setExplosion(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   */

  @DISPID(261)
  @PropGet
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(261)
  @PropPut
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaLocal"
   * </p>
   */

  @DISPID(263)
  @PropGet
  java.lang.String getFormulaLocal();


  /**
   * <p>
   * Setter method for the COM property "FormulaLocal"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(263)
  @PropPut
  void setFormulaLocal(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1"
   * </p>
   */

  @DISPID(264)
  @PropGet
  java.lang.String getFormulaR1C1();


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(264)
  @PropPut
  void setFormulaR1C1(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FormulaR1C1Local"
   * </p>
   */

  @DISPID(265)
  @PropGet
  java.lang.String getFormulaR1C1Local();


  /**
   * <p>
   * Setter method for the COM property "FormulaR1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(265)
  @PropPut
  void setFormulaR1C1Local(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDataLabels"
   * </p>
   */

  @DISPID(78)
  @PropGet
  boolean getHasDataLabels();


  /**
   * <p>
   * Setter method for the COM property "HasDataLabels"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(78)
  @PropPut
  void setHasDataLabels(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasErrorBars"
   * </p>
   */

  @DISPID(160)
  @PropGet
  boolean getHasErrorBars();


  /**
   * <p>
   * Setter method for the COM property "HasErrorBars"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(160)
  @PropPut
  void setHasErrorBars(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   */

  @DISPID(129)
  @PropGet
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   */

  @DISPID(1663)
  @PropGet
  com.exceljava.com4j.excel.ChartFillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "InvertIfNegative"
   * </p>
   */

  @DISPID(132)
  @PropGet
  boolean getInvertIfNegative();


  /**
   * <p>
   * Setter method for the COM property "InvertIfNegative"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(132)
  @PropPut
  void setInvertIfNegative(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerBackgroundColor"
   * </p>
   */

  @DISPID(73)
  @PropGet
  int getMarkerBackgroundColor();


  /**
   * <p>
   * Setter method for the COM property "MarkerBackgroundColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(73)
  @PropPut
  void setMarkerBackgroundColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerBackgroundColorIndex"
   * </p>
   */

  @DISPID(74)
  @PropGet
  com.exceljava.com4j.excel.XlColorIndex getMarkerBackgroundColorIndex();


  /**
   * <p>
   * Setter method for the COM property "MarkerBackgroundColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @DISPID(74)
  @PropPut
  void setMarkerBackgroundColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerForegroundColor"
   * </p>
   */

  @DISPID(75)
  @PropGet
  int getMarkerForegroundColor();


  /**
   * <p>
   * Setter method for the COM property "MarkerForegroundColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(75)
  @PropPut
  void setMarkerForegroundColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerForegroundColorIndex"
   * </p>
   */

  @DISPID(76)
  @PropGet
  com.exceljava.com4j.excel.XlColorIndex getMarkerForegroundColorIndex();


  /**
   * <p>
   * Setter method for the COM property "MarkerForegroundColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @DISPID(76)
  @PropPut
  void setMarkerForegroundColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerSize"
   * </p>
   */

  @DISPID(231)
  @PropGet
  int getMarkerSize();


  /**
   * <p>
   * Setter method for the COM property "MarkerSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(231)
  @PropPut
  void setMarkerSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerStyle"
   * </p>
   */

  @DISPID(72)
  @PropGet
  com.exceljava.com4j.excel.XlMarkerStyle getMarkerStyle();


  /**
   * <p>
   * Setter method for the COM property "MarkerStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlMarkerStyle parameter.
   */

  @DISPID(72)
  @PropPut
  void setMarkerStyle(
    com.exceljava.com4j.excel.XlMarkerStyle rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   */

  @DISPID(211)
  java.lang.Object paste();


  /**
   * <p>
   * Getter method for the COM property "PictureType"
   * </p>
   */

  @DISPID(161)
  @PropGet
  com.exceljava.com4j.excel.XlChartPictureType getPictureType();


  /**
   * <p>
   * Setter method for the COM property "PictureType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartPictureType parameter.
   */

  @DISPID(161)
  @PropPut
  void setPictureType(
    com.exceljava.com4j.excel.XlChartPictureType rhs);


  /**
   * <p>
   * Getter method for the COM property "PictureUnit"
   * </p>
   */

  @DISPID(162)
  @PropGet
  int getPictureUnit();


  /**
   * <p>
   * Setter method for the COM property "PictureUnit"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(162)
  @PropPut
  void setPictureUnit(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "PlotOrder"
   * </p>
   */

  @DISPID(228)
  @PropGet
  int getPlotOrder();


  /**
   * <p>
   * Setter method for the COM property "PlotOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(228)
  @PropPut
  void setPlotOrder(
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * points(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(70)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject points();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(70)
  com4j.Com4jObject points(
    @Optional java.lang.Object index);


  /**
   */

  @DISPID(235)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "Smooth"
   * </p>
   */

  @DISPID(163)
  @PropGet
  boolean getSmooth();


  /**
   * <p>
   * Setter method for the COM property "Smooth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(163)
  @PropPut
  void setSmooth(
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
   * trendlines(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(154)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject trendlines();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(154)
  com4j.Com4jObject trendlines(
    @Optional java.lang.Object index);


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
   * Getter method for the COM property "ChartType"
   * </p>
   */

  @DISPID(1400)
  @PropGet
  com.exceljava.com4j.excel.XlChartType getChartType();


  /**
   * <p>
   * Setter method for the COM property "ChartType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartType parameter.
   */

  @DISPID(1400)
  @PropPut
  void setChartType(
    com.exceljava.com4j.excel.XlChartType rhs);


  /**
   * @param chartType Mandatory com.exceljava.com4j.excel.XlChartType parameter.
   */

  @DISPID(1401)
  void applyCustomType(
    com.exceljava.com4j.excel.XlChartType chartType);


  /**
   * <p>
   * Getter method for the COM property "Values"
   * </p>
   */

  @DISPID(164)
  @PropGet
  java.lang.Object getValues();


  /**
   * <p>
   * Setter method for the COM property "Values"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(164)
  @PropPut
  void setValues(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "XValues"
   * </p>
   */

  @DISPID(1111)
  @PropGet
  java.lang.Object getXValues();


  /**
   * <p>
   * Setter method for the COM property "XValues"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1111)
  @PropPut
  void setXValues(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BubbleSizes"
   * </p>
   */

  @DISPID(1664)
  @PropGet
  java.lang.Object getBubbleSizes();


  /**
   * <p>
   * Setter method for the COM property "BubbleSizes"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1664)
  @PropPut
  void setBubbleSizes(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BarShape"
   * </p>
   */

  @DISPID(1403)
  @PropGet
  com.exceljava.com4j.excel.XlBarShape getBarShape();


  /**
   * <p>
   * Setter method for the COM property "BarShape"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlBarShape parameter.
   */

  @DISPID(1403)
  @PropPut
  void setBarShape(
    com.exceljava.com4j.excel.XlBarShape rhs);


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToSides"
   * </p>
   */

  @DISPID(1659)
  @PropGet
  boolean getApplyPictToSides();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToSides"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1659)
  @PropPut
  void setApplyPictToSides(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToFront"
   * </p>
   */

  @DISPID(1660)
  @PropGet
  boolean getApplyPictToFront();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToFront"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1660)
  @PropPut
  void setApplyPictToFront(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToEnd"
   * </p>
   */

  @DISPID(1661)
  @PropGet
  boolean getApplyPictToEnd();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToEnd"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1661)
  @PropPut
  void setApplyPictToEnd(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Has3DEffect"
   * </p>
   */

  @DISPID(1665)
  @PropGet
  boolean getHas3DEffect();


  /**
   * <p>
   * Setter method for the COM property "Has3DEffect"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1665)
  @PropPut
  void setHas3DEffect(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   */

  @DISPID(103)
  @PropGet
  boolean getShadow();


  /**
   * <p>
   * Setter method for the COM property "Shadow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(103)
  @PropPut
  void setShadow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasLeaderLines"
   * </p>
   */

  @DISPID(1394)
  @PropGet
  boolean getHasLeaderLines();


  /**
   * <p>
   * Setter method for the COM property "HasLeaderLines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1394)
  @PropPut
  void setHasLeaderLines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LeaderLines"
   * </p>
   */

  @DISPID(1666)
  @PropGet
  com.exceljava.com4j.excel.LeaderLines getLeaderLines();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName,
    @Optional java.lang.Object showCategoryName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName,
    @Optional java.lang.Object showCategoryName,
    @Optional java.lang.Object showValue);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName,
    @Optional java.lang.Object showCategoryName,
    @Optional java.lang.Object showValue,
    @Optional java.lang.Object showPercentage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName,
    @Optional java.lang.Object showCategoryName,
    @Optional java.lang.Object showValue,
    @Optional java.lang.Object showPercentage,
    @Optional java.lang.Object showBubbleSize);

  /**
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @param separator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional java.lang.Object legendKey,
    @Optional java.lang.Object autoText,
    @Optional java.lang.Object hasLeaderLines,
    @Optional java.lang.Object showSeriesName,
    @Optional java.lang.Object showCategoryName,
    @Optional java.lang.Object showValue,
    @Optional java.lang.Object showPercentage,
    @Optional java.lang.Object showBubbleSize,
    @Optional java.lang.Object separator);


  /**
   * <p>
   * Getter method for the COM property "PictureUnit2"
   * </p>
   */

  @DISPID(2649)
  @PropGet
  double getPictureUnit2();


  /**
   * <p>
   * Setter method for the COM property "PictureUnit2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2649)
  @PropPut
  void setPictureUnit2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   */

  @DISPID(116)
  @PropGet
  com.exceljava.com4j.excel.ChartFormat getFormat();


  /**
   * <p>
   * Getter method for the COM property "PlotColorIndex"
   * </p>
   */

  @DISPID(2915)
  @PropGet
  int getPlotColorIndex();


  /**
   * <p>
   * Getter method for the COM property "InvertColor"
   * </p>
   */

  @DISPID(2916)
  @PropGet
  int getInvertColor();


  /**
   * <p>
   * Setter method for the COM property "InvertColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2916)
  @PropPut
  void setInvertColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "InvertColorIndex"
   * </p>
   */

  @DISPID(2917)
  @PropGet
  int getInvertColorIndex();


  /**
   * <p>
   * Setter method for the COM property "InvertColorIndex"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2917)
  @PropPut
  void setInvertColorIndex(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "IsFiltered"
   * </p>
   */

  @DISPID(3083)
  @PropGet
  boolean getIsFiltered();


  /**
   * <p>
   * Setter method for the COM property "IsFiltered"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3083)
  @PropPut
  void setIsFiltered(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ParentDataLabelOption"
   * </p>
   */

  @DISPID(3204)
  @PropGet
  com.exceljava.com4j.excel.XlParentDataLabelOptions getParentDataLabelOption();


  /**
   * <p>
   * Setter method for the COM property "ParentDataLabelOption"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlParentDataLabelOptions parameter.
   */

  @DISPID(3204)
  @PropPut
  void setParentDataLabelOption(
    com.exceljava.com4j.excel.XlParentDataLabelOptions rhs);


  /**
   * <p>
   * Getter method for the COM property "QuartileCalculationInclusiveMedian"
   * </p>
   */

  @DISPID(3205)
  @PropGet
  boolean getQuartileCalculationInclusiveMedian();


  /**
   * <p>
   * Setter method for the COM property "QuartileCalculationInclusiveMedian"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3205)
  @PropPut
  void setQuartileCalculationInclusiveMedian(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ValueSortOrder"
   * </p>
   */

  @DISPID(3229)
  @PropGet
  com.exceljava.com4j.excel.XlValueSortOrder getValueSortOrder();


  /**
   * <p>
   * Setter method for the COM property "ValueSortOrder"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlValueSortOrder parameter.
   */

  @DISPID(3229)
  @PropPut
  void setValueSortOrder(
    com.exceljava.com4j.excel.XlValueSortOrder rhs);


  /**
   * <p>
   * Getter method for the COM property "GeoProjectionType"
   * </p>
   */

  @DISPID(3250)
  @PropGet
  com.exceljava.com4j.excel.XlGeoProjectionType getGeoProjectionType();


  /**
   * <p>
   * Setter method for the COM property "GeoProjectionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlGeoProjectionType parameter.
   */

  @DISPID(3250)
  @PropPut
  void setGeoProjectionType(
    com.exceljava.com4j.excel.XlGeoProjectionType rhs);


  /**
   * <p>
   * Getter method for the COM property "GeoMappingLevel"
   * </p>
   */

  @DISPID(3251)
  @PropGet
  com.exceljava.com4j.excel.XlGeoMappingLevel getGeoMappingLevel();


  /**
   * <p>
   * Setter method for the COM property "GeoMappingLevel"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlGeoMappingLevel parameter.
   */

  @DISPID(3251)
  @PropPut
  void setGeoMappingLevel(
    com.exceljava.com4j.excel.XlGeoMappingLevel rhs);


  /**
   * <p>
   * Getter method for the COM property "RegionLabelOption"
   * </p>
   */

  @DISPID(3252)
  @PropGet
  com.exceljava.com4j.excel.XlRegionLabelOptions getRegionLabelOption();


  /**
   * <p>
   * Setter method for the COM property "RegionLabelOption"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRegionLabelOptions parameter.
   */

  @DISPID(3252)
  @PropPut
  void setRegionLabelOption(
    com.exceljava.com4j.excel.XlRegionLabelOptions rhs);


  /**
   * <p>
   * Getter method for the COM property "SeriesColorGradientStyle"
   * </p>
   */

  @DISPID(3261)
  @PropGet
  com.exceljava.com4j.excel.XlSeriesColorGradientStyle getSeriesColorGradientStyle();


  /**
   * <p>
   * Setter method for the COM property "SeriesColorGradientStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSeriesColorGradientStyle parameter.
   */

  @DISPID(3261)
  @PropPut
  void setSeriesColorGradientStyle(
    com.exceljava.com4j.excel.XlSeriesColorGradientStyle rhs);


  /**
   * <p>
   * Getter method for the COM property "SeriesColorMinGradientStop"
   * </p>
   */

  @DISPID(3262)
  @PropGet
  com.exceljava.com4j.excel.ChartSeriesGradientStopData getSeriesColorMinGradientStop();


  /**
   * <p>
   * Getter method for the COM property "SeriesColorMidGradientStop"
   * </p>
   */

  @DISPID(3263)
  @PropGet
  com.exceljava.com4j.excel.ChartSeriesGradientStopData getSeriesColorMidGradientStop();


  /**
   * <p>
   * Getter method for the COM property "SeriesColorMaxGradientStop"
   * </p>
   */

  @DISPID(3264)
  @PropGet
  com.exceljava.com4j.excel.ChartSeriesGradientStopData getSeriesColorMaxGradientStop();


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @DISPID(3253)
  void setProperty(
    java.lang.String id,
    java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(3256)
  java.lang.Object getProperty(
    java.lang.String id);


  // Properties:
}
