package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface DataLabel extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   */

  @DISPID(235)
  java.lang.Object select();


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

  @DISPID(117)
  java.lang.Object delete();


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
   * Getter method for the COM property "Caption"
   * </p>
   */

  @DISPID(139)
  @PropGet
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(139)
  @PropPut
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(603)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters getCharacters();

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional java.lang.Object start);

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(603)
  @PropGet
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional java.lang.Object start,
    @Optional java.lang.Object length);


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   */

  @DISPID(146)
  @PropGet
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   */

  @DISPID(136)
  @PropGet
  java.lang.Object getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(136)
  @PropPut
  void setHorizontalAlignment(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(127)
  @PropPut
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  java.lang.Object getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    java.lang.Object rhs);


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
   * Getter method for the COM property "Text"
   * </p>
   */

  @DISPID(138)
  @PropGet
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(138)
  @PropPut
  void setText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(126)
  @PropPut
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   */

  @DISPID(137)
  @PropGet
  java.lang.Object getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(137)
  @PropPut
  void setVerticalAlignment(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   */

  @DISPID(975)
  @PropGet
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(975)
  @PropPut
  void setReadingOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoScaleFont"
   * </p>
   */

  @DISPID(1525)
  @PropGet
  java.lang.Object getAutoScaleFont();


  /**
   * <p>
   * Setter method for the COM property "AutoScaleFont"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1525)
  @PropPut
  void setAutoScaleFont(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoText"
   * </p>
   */

  @DISPID(135)
  @PropGet
  boolean getAutoText();


  /**
   * <p>
   * Setter method for the COM property "AutoText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(135)
  @PropPut
  void setAutoText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   */

  @DISPID(193)
  @PropGet
  java.lang.String getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(193)
  @PropPut
  void setNumberFormat(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLinked"
   * </p>
   */

  @DISPID(194)
  @PropGet
  boolean getNumberFormatLinked();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLinked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(194)
  @PropPut
  void setNumberFormatLinked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   */

  @DISPID(1097)
  @PropGet
  java.lang.Object getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1097)
  @PropPut
  void setNumberFormatLocal(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowLegendKey"
   * </p>
   */

  @DISPID(171)
  @PropGet
  boolean getShowLegendKey();


  /**
   * <p>
   * Setter method for the COM property "ShowLegendKey"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(171)
  @PropPut
  void setShowLegendKey(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  java.lang.Object getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(108)
  @PropPut
  void setType(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  com.exceljava.com4j.excel.XlDataLabelPosition getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDataLabelPosition parameter.
   */

  @DISPID(133)
  @PropPut
  void setPosition(
    com.exceljava.com4j.excel.XlDataLabelPosition rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowSeriesName"
   * </p>
   */

  @DISPID(2022)
  @PropGet
  boolean getShowSeriesName();


  /**
   * <p>
   * Setter method for the COM property "ShowSeriesName"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2022)
  @PropPut
  void setShowSeriesName(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowCategoryName"
   * </p>
   */

  @DISPID(2023)
  @PropGet
  boolean getShowCategoryName();


  /**
   * <p>
   * Setter method for the COM property "ShowCategoryName"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2023)
  @PropPut
  void setShowCategoryName(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowValue"
   * </p>
   */

  @DISPID(2024)
  @PropGet
  boolean getShowValue();


  /**
   * <p>
   * Setter method for the COM property "ShowValue"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2024)
  @PropPut
  void setShowValue(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowPercentage"
   * </p>
   */

  @DISPID(2025)
  @PropGet
  boolean getShowPercentage();


  /**
   * <p>
   * Setter method for the COM property "ShowPercentage"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2025)
  @PropPut
  void setShowPercentage(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowBubbleSize"
   * </p>
   */

  @DISPID(2026)
  @PropGet
  boolean getShowBubbleSize();


  /**
   * <p>
   * Setter method for the COM property "ShowBubbleSize"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2026)
  @PropPut
  void setShowBubbleSize(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Separator"
   * </p>
   */

  @DISPID(2027)
  @PropGet
  java.lang.Object getSeparator();


  /**
   * <p>
   * Setter method for the COM property "Separator"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2027)
  @PropPut
  void setSeparator(
    java.lang.Object rhs);


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
   * Getter method for the COM property "_Height"
   * </p>
   */

  @DISPID(3084)
  @PropGet
  double get_Height();


  /**
   * <p>
   * Getter method for the COM property "_Width"
   * </p>
   */

  @DISPID(3085)
  @PropGet
  double get_Width();


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
   * Getter method for the COM property "ShowRange"
   * </p>
   */

  @DISPID(3086)
  @PropGet
  boolean getShowRange();


  /**
   * <p>
   * Setter method for the COM property "ShowRange"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3086)
  @PropPut
  void setShowRange(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    double rhs);


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
