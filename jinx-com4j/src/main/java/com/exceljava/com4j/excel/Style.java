package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Style extends Com4jObject {
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
   * Getter method for the COM property "AddIndent"
   * </p>
   */

  @DISPID(1063)
  @PropGet
  boolean getAddIndent();


  /**
   * <p>
   * Setter method for the COM property "AddIndent"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1063)
  @PropPut
  void setAddIndent(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BuiltIn"
   * </p>
   */

  @DISPID(553)
  @PropGet
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   */

  @DISPID(435)
  @PropGet
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   */

  @DISPID(117)
  java.lang.Object delete();


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
   * Getter method for the COM property "FormulaHidden"
   * </p>
   */

  @DISPID(262)
  @PropGet
  boolean getFormulaHidden();


  /**
   * <p>
   * Setter method for the COM property "FormulaHidden"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(262)
  @PropPut
  void setFormulaHidden(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   */

  @DISPID(136)
  @PropGet
  com.exceljava.com4j.excel.XlHAlign getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlHAlign parameter.
   */

  @DISPID(136)
  @PropPut
  void setHorizontalAlignment(
    com.exceljava.com4j.excel.XlHAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeAlignment"
   * </p>
   */

  @DISPID(413)
  @PropGet
  boolean getIncludeAlignment();


  /**
   * <p>
   * Setter method for the COM property "IncludeAlignment"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(413)
  @PropPut
  void setIncludeAlignment(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeBorder"
   * </p>
   */

  @DISPID(414)
  @PropGet
  boolean getIncludeBorder();


  /**
   * <p>
   * Setter method for the COM property "IncludeBorder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(414)
  @PropPut
  void setIncludeBorder(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeFont"
   * </p>
   */

  @DISPID(415)
  @PropGet
  boolean getIncludeFont();


  /**
   * <p>
   * Setter method for the COM property "IncludeFont"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(415)
  @PropPut
  void setIncludeFont(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeNumber"
   * </p>
   */

  @DISPID(416)
  @PropGet
  boolean getIncludeNumber();


  /**
   * <p>
   * Setter method for the COM property "IncludeNumber"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(416)
  @PropPut
  void setIncludeNumber(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludePatterns"
   * </p>
   */

  @DISPID(417)
  @PropGet
  boolean getIncludePatterns();


  /**
   * <p>
   * Setter method for the COM property "IncludePatterns"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(417)
  @PropPut
  void setIncludePatterns(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeProtection"
   * </p>
   */

  @DISPID(418)
  @PropGet
  boolean getIncludeProtection();


  /**
   * <p>
   * Setter method for the COM property "IncludeProtection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(418)
  @PropPut
  void setIncludeProtection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IndentLevel"
   * </p>
   */

  @DISPID(201)
  @PropGet
  int getIndentLevel();


  /**
   * <p>
   * Setter method for the COM property "IndentLevel"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(201)
  @PropPut
  void setIndentLevel(
    int rhs);


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
   * Getter method for the COM property "Locked"
   * </p>
   */

  @DISPID(269)
  @PropGet
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(269)
  @PropPut
  void setLocked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MergeCells"
   * </p>
   */

  @DISPID(208)
  @PropGet
  java.lang.Object getMergeCells();


  /**
   * <p>
   * Setter method for the COM property "MergeCells"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(208)
  @PropPut
  void setMergeCells(
    java.lang.Object rhs);


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
   * Getter method for the COM property "NameLocal"
   * </p>
   */

  @DISPID(937)
  @PropGet
  java.lang.String getNameLocal();


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
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   */

  @DISPID(1097)
  @PropGet
  java.lang.String getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1097)
  @PropPut
  void setNumberFormatLocal(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  com.exceljava.com4j.excel.XlOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOrientation parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    com.exceljava.com4j.excel.XlOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "ShrinkToFit"
   * </p>
   */

  @DISPID(209)
  @PropGet
  boolean getShrinkToFit();


  /**
   * <p>
   * Setter method for the COM property "ShrinkToFit"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(209)
  @PropPut
  void setShrinkToFit(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.String getValue();


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   */

  @DISPID(137)
  @PropGet
  com.exceljava.com4j.excel.XlVAlign getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlVAlign parameter.
   */

  @DISPID(137)
  @PropPut
  void setVerticalAlignment(
    com.exceljava.com4j.excel.XlVAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "WrapText"
   * </p>
   */

  @DISPID(276)
  @PropGet
  boolean getWrapText();


  /**
   * <p>
   * Setter method for the COM property "WrapText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(276)
  @PropPut
  void setWrapText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.String get_Default();


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


  // Properties:
}
