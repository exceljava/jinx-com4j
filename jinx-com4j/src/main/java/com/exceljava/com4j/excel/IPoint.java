package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002086A-0001-0000-C000-000000000046}")
public interface IPoint extends Com4jObject {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

  /**
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Border
   */

  @VTID(11)
  com.exceljava.com4j.excel.Border getBorder();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object clearFormats();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copy();


  /**
   * <p>
   * Getter method for the COM property "DataLabel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DataLabel
   */

  @VTID(14)
  com.exceljava.com4j.excel.DataLabel getDataLabel();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "Explosion"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getExplosion();


  /**
   * <p>
   * Setter method for the COM property "Explosion"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(17)
  void setExplosion(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDataLabel"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getHasDataLabel();


  /**
   * <p>
   * Setter method for the COM property "HasDataLabel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setHasDataLabel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(20)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "InvertIfNegative"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(21)
  boolean getInvertIfNegative();


  /**
   * <p>
   * Setter method for the COM property "InvertIfNegative"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(22)
  void setInvertIfNegative(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerBackgroundColor"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(23)
  int getMarkerBackgroundColor();


  /**
   * <p>
   * Setter method for the COM property "MarkerBackgroundColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(24)
  void setMarkerBackgroundColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerBackgroundColorIndex"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlColorIndex
   */

  @VTID(25)
  com.exceljava.com4j.excel.XlColorIndex getMarkerBackgroundColorIndex();


  /**
   * <p>
   * Setter method for the COM property "MarkerBackgroundColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @VTID(26)
  void setMarkerBackgroundColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerForegroundColor"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(27)
  int getMarkerForegroundColor();


  /**
   * <p>
   * Setter method for the COM property "MarkerForegroundColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(28)
  void setMarkerForegroundColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerForegroundColorIndex"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlColorIndex
   */

  @VTID(29)
  com.exceljava.com4j.excel.XlColorIndex getMarkerForegroundColorIndex();


  /**
   * <p>
   * Setter method for the COM property "MarkerForegroundColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @VTID(30)
  void setMarkerForegroundColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerSize"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(31)
  int getMarkerSize();


  /**
   * <p>
   * Setter method for the COM property "MarkerSize"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(32)
  void setMarkerSize(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MarkerStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlMarkerStyle
   */

  @VTID(33)
  com.exceljava.com4j.excel.XlMarkerStyle getMarkerStyle();


  /**
   * <p>
   * Setter method for the COM property "MarkerStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlMarkerStyle parameter.
   */

  @VTID(34)
  void setMarkerStyle(
    com.exceljava.com4j.excel.XlMarkerStyle rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(35)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object paste();


  /**
   * <p>
   * Getter method for the COM property "PictureType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlChartPictureType
   */

  @VTID(36)
  com.exceljava.com4j.excel.XlChartPictureType getPictureType();


  /**
   * <p>
   * Setter method for the COM property "PictureType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartPictureType parameter.
   */

  @VTID(37)
  void setPictureType(
    com.exceljava.com4j.excel.XlChartPictureType rhs);


  /**
   * <p>
   * Getter method for the COM property "PictureUnit"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(38)
  int getPictureUnit();


  /**
   * <p>
   * Setter method for the COM property "PictureUnit"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(39)
  void setPictureUnit(
    int rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(40)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToSides"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(41)
  boolean getApplyPictToSides();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToSides"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(42)
  void setApplyPictToSides(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToFront"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(43)
  boolean getApplyPictToFront();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToFront"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(44)
  void setApplyPictToFront(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ApplyPictToEnd"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(45)
  boolean getApplyPictToEnd();


  /**
   * <p>
   * Setter method for the COM property "ApplyPictToEnd"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(46)
  void setApplyPictToEnd(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(47)
  boolean getShadow();


  /**
   * <p>
   * Setter method for the COM property "Shadow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(48)
  void setShadow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SecondaryPlot"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(49)
  boolean getSecondaryPlot();


  /**
   * <p>
   * Setter method for the COM property "SecondaryPlot"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(50)
  void setSecondaryPlot(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartFillFormat
   */

  @VTID(51)
  com.exceljava.com4j.excel.ChartFillFormat getFill();


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {10}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 10}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 10}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 10}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 10}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 10}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 10}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 10}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 10}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 10}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=10)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object separator);


  /**
   * <p>
   * Getter method for the COM property "Has3DEffect"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(53)
  boolean getHas3DEffect();


  /**
   * <p>
   * Setter method for the COM property "Has3DEffect"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(54)
  void setHas3DEffect(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PictureUnit2"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(55)
  double getPictureUnit2();


  /**
   * <p>
   * Setter method for the COM property "PictureUnit2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(56)
  void setPictureUnit2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartFormat
   */

  @VTID(57)
  com.exceljava.com4j.excel.ChartFormat getFormat();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(58)
  double getHeight();


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(59)
  double getWidth();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(60)
  double getTop();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(61)
  double getLeft();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(62)
  java.lang.String getName();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPieSliceIndex parameter index is set to 2</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pieSliceLocation(loc, 2);
   * </code>
   * </p>
   * @param loc Mandatory com.exceljava.com4j.excel.XlPieSliceLocation parameter.
   * @return  Returns a value of type double
   */

  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlPieSliceIndex.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"2"})
  @ReturnValue(index=2)
  double pieSliceLocation(
    com.exceljava.com4j.excel.XlPieSliceLocation loc);

  /**
   * @param loc Mandatory com.exceljava.com4j.excel.XlPieSliceLocation parameter.
   * @param index Optional parameter. Default value is 2
   * @return  Returns a value of type double
   */

  @VTID(63)
  double pieSliceLocation(
    com.exceljava.com4j.excel.XlPieSliceLocation loc,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPieSliceIndex index);


  /**
   * <p>
   * Getter method for the COM property "IsTotal"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(64)
  boolean getIsTotal();


  /**
   * <p>
   * Setter method for the COM property "IsTotal"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(65)
  void setIsTotal(
    boolean rhs);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @VTID(66)
  void setProperty(
    java.lang.String id,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(67)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String id);


  // Properties:
}
