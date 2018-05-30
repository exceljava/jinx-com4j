package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020852-0001-0000-C000-000000000046}")
public interface IStyle extends Com4jObject {
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
   * Getter method for the COM property "AddIndent"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getAddIndent();


  /**
   * <p>
   * Setter method for the COM property "AddIndent"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setAddIndent(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BuiltIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Borders
   */

  @VTID(13)
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(15)
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "FormulaHidden"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getFormulaHidden();


  /**
   * <p>
   * Setter method for the COM property "FormulaHidden"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setFormulaHidden(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalAlignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlHAlign
   */

  @VTID(18)
  com.exceljava.com4j.excel.XlHAlign getHorizontalAlignment();


  /**
   * <p>
   * Setter method for the COM property "HorizontalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlHAlign parameter.
   */

  @VTID(19)
  void setHorizontalAlignment(
    com.exceljava.com4j.excel.XlHAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeAlignment"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getIncludeAlignment();


  /**
   * <p>
   * Setter method for the COM property "IncludeAlignment"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setIncludeAlignment(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeBorder"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getIncludeBorder();


  /**
   * <p>
   * Setter method for the COM property "IncludeBorder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setIncludeBorder(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeFont"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getIncludeFont();


  /**
   * <p>
   * Setter method for the COM property "IncludeFont"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setIncludeFont(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeNumber"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getIncludeNumber();


  /**
   * <p>
   * Setter method for the COM property "IncludeNumber"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setIncludeNumber(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludePatterns"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getIncludePatterns();


  /**
   * <p>
   * Setter method for the COM property "IncludePatterns"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setIncludePatterns(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeProtection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getIncludeProtection();


  /**
   * <p>
   * Setter method for the COM property "IncludeProtection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setIncludeProtection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IndentLevel"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(32)
  int getIndentLevel();


  /**
   * <p>
   * Setter method for the COM property "IndentLevel"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(33)
  void setIndentLevel(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(34)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(35)
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(36)
  void setLocked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MergeCells"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(37)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getMergeCells();


  /**
   * <p>
   * Setter method for the COM property "MergeCells"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(38)
  void setMergeCells(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getName(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getName();

  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(39)
  java.lang.String getName(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "NameLocal"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(40)
  java.lang.String getNameLocal();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(41)
  java.lang.String getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(42)
  void setNumberFormat(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormatLocal"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(43)
  java.lang.String getNumberFormatLocal();


  /**
   * <p>
   * Setter method for the COM property "NumberFormatLocal"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(44)
  void setNumberFormatLocal(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlOrientation
   */

  @VTID(45)
  com.exceljava.com4j.excel.XlOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOrientation parameter.
   */

  @VTID(46)
  void setOrientation(
    com.exceljava.com4j.excel.XlOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "ShrinkToFit"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(47)
  boolean getShrinkToFit();


  /**
   * <p>
   * Setter method for the COM property "ShrinkToFit"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(48)
  void setShrinkToFit(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getValue(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getValue();

  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(49)
  java.lang.String getValue(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "VerticalAlignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlVAlign
   */

  @VTID(50)
  com.exceljava.com4j.excel.XlVAlign getVerticalAlignment();


  /**
   * <p>
   * Setter method for the COM property "VerticalAlignment"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlVAlign parameter.
   */

  @VTID(51)
  void setVerticalAlignment(
    com.exceljava.com4j.excel.XlVAlign rhs);


  /**
   * <p>
   * Getter method for the COM property "WrapText"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(52)
  boolean getWrapText();


  /**
   * <p>
   * Setter method for the COM property "WrapText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(53)
  void setWrapText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_Default(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(54)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String get_Default();

  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(54)
  @DefaultMethod
  java.lang.String get_Default(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ReadingOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(55)
  int getReadingOrder();


  /**
   * <p>
   * Setter method for the COM property "ReadingOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(56)
  void setReadingOrder(
    int rhs);


  // Properties:
}
