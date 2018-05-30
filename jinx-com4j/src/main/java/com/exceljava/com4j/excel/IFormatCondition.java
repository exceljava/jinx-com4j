package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024425-0001-0000-C000-000000000046}")
public interface IFormatCondition extends Com4jObject {
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
     * <li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Modify(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _Modify(
    com.exceljava.com4j.excel.XlFormatConditionType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Modify(type, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _Modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Modify(type, operator, formula1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _Modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  void _Modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getType();


  /**
   * <p>
   * Getter method for the COM property "Operator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getOperator();


  /**
   * <p>
   * Getter method for the COM property "Formula1"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getFormula1();


  /**
   * <p>
   * Getter method for the COM property "Formula2"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getFormula2();


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(15)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Borders
   */

  @VTID(16)
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(17)
  com.exceljava.com4j.excel.Font getFont();


  /**
   */

  @VTID(18)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, operator, formula1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, operator, formula1, formula2, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter operator2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, operator, formula1, formula2, string, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
  void modify(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator2);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(20)
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(21)
  void setText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextOperator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlContainsOperator
   */

  @VTID(22)
  com.exceljava.com4j.excel.XlContainsOperator getTextOperator();


  /**
   * <p>
   * Setter method for the COM property "TextOperator"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlContainsOperator parameter.
   */

  @VTID(23)
  void setTextOperator(
    com.exceljava.com4j.excel.XlContainsOperator rhs);


  /**
   * <p>
   * Getter method for the COM property "DateOperator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTimePeriods
   */

  @VTID(24)
  com.exceljava.com4j.excel.XlTimePeriods getDateOperator();


  /**
   * <p>
   * Setter method for the COM property "DateOperator"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTimePeriods parameter.
   */

  @VTID(25)
  void setDateOperator(
    com.exceljava.com4j.excel.XlTimePeriods rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(26)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(27)
  void setNumberFormat(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Priority"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(28)
  int getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(29)
  void setPriority(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "StopIfTrue"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getStopIfTrue();


  /**
   * <p>
   * Setter method for the COM property "StopIfTrue"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setStopIfTrue(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AppliesTo"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(32)
  com.exceljava.com4j.excel.Range getAppliesTo();


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(33)
  void modifyAppliesToRange(
    com.exceljava.com4j.excel.Range range);


  /**
   */

  @VTID(34)
  void setFirstPriority();


  /**
   */

  @VTID(35)
  void setLastPriority();


  /**
   * <p>
   * Getter method for the COM property "PTCondition"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(36)
  boolean getPTCondition();


  /**
   * <p>
   * Getter method for the COM property "ScopeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotConditionScope
   */

  @VTID(37)
  com.exceljava.com4j.excel.XlPivotConditionScope getScopeType();


  /**
   * <p>
   * Setter method for the COM property "ScopeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotConditionScope parameter.
   */

  @VTID(38)
  void setScopeType(
    com.exceljava.com4j.excel.XlPivotConditionScope rhs);


  // Properties:
}
