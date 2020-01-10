package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface AutoCorrect extends Com4jObject {
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
   * @param what Mandatory java.lang.String parameter.
   * @param replacement Mandatory java.lang.String parameter.
   */

  @DISPID(1146)
  java.lang.Object addReplacement(
    java.lang.String what,
    java.lang.String replacement);


  /**
   * <p>
   * Getter method for the COM property "CapitalizeNamesOfDays"
   * </p>
   */

  @DISPID(1150)
  @PropGet
  boolean getCapitalizeNamesOfDays();


  /**
   * <p>
   * Setter method for the COM property "CapitalizeNamesOfDays"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1150)
  @PropPut
  void setCapitalizeNamesOfDays(
    boolean rhs);


  /**
   * @param what Mandatory java.lang.String parameter.
   */

  @DISPID(1147)
  java.lang.Object deleteReplacement(
    java.lang.String what);


  /**
   * <p>
   * Getter method for the COM property "ReplacementList"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getReplacementList(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1151)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getReplacementList();

  /**
   * <p>
   * Getter method for the COM property "ReplacementList"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1151)
  @PropGet
  java.lang.Object getReplacementList(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "ReplacementList"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setReplacementList(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1151)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setReplacementList(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "ReplacementList"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1151)
  @PropPut
  void setReplacementList(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ReplaceText"
   * </p>
   */

  @DISPID(1148)
  @PropGet
  boolean getReplaceText();


  /**
   * <p>
   * Setter method for the COM property "ReplaceText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1148)
  @PropPut
  void setReplaceText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TwoInitialCapitals"
   * </p>
   */

  @DISPID(1149)
  @PropGet
  boolean getTwoInitialCapitals();


  /**
   * <p>
   * Setter method for the COM property "TwoInitialCapitals"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1149)
  @PropPut
  void setTwoInitialCapitals(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CorrectSentenceCap"
   * </p>
   */

  @DISPID(1619)
  @PropGet
  boolean getCorrectSentenceCap();


  /**
   * <p>
   * Setter method for the COM property "CorrectSentenceCap"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1619)
  @PropPut
  void setCorrectSentenceCap(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CorrectCapsLock"
   * </p>
   */

  @DISPID(1620)
  @PropGet
  boolean getCorrectCapsLock();


  /**
   * <p>
   * Setter method for the COM property "CorrectCapsLock"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1620)
  @PropPut
  void setCorrectCapsLock(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAutoCorrectOptions"
   * </p>
   */

  @DISPID(1926)
  @PropGet
  boolean getDisplayAutoCorrectOptions();


  /**
   * <p>
   * Setter method for the COM property "DisplayAutoCorrectOptions"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1926)
  @PropPut
  void setDisplayAutoCorrectOptions(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoExpandListRange"
   * </p>
   */

  @DISPID(2294)
  @PropGet
  boolean getAutoExpandListRange();


  /**
   * <p>
   * Setter method for the COM property "AutoExpandListRange"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2294)
  @PropPut
  void setAutoExpandListRange(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFillFormulasInLists"
   * </p>
   */

  @DISPID(2642)
  @PropGet
  boolean getAutoFillFormulasInLists();


  /**
   * <p>
   * Setter method for the COM property "AutoFillFormulasInLists"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2642)
  @PropPut
  void setAutoFillFormulasInLists(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "KeepGeneralFormatLeadingZerosAsText"
   * </p>
   */

  @DISPID(3348)
  @PropGet
  boolean getKeepGeneralFormatLeadingZerosAsText();


  /**
   * <p>
   * Setter method for the COM property "KeepGeneralFormatLeadingZerosAsText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3348)
  @PropPut
  void setKeepGeneralFormatLeadingZerosAsText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "KeepGeneralFormatLargeNumbersAsText"
   * </p>
   */

  @DISPID(3349)
  @PropGet
  boolean getKeepGeneralFormatLargeNumbersAsText();


  /**
   * <p>
   * Setter method for the COM property "KeepGeneralFormatLargeNumbersAsText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3349)
  @PropPut
  void setKeepGeneralFormatLargeNumbersAsText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "KeepGeneralFormatDigitsWithEAsText"
   * </p>
   */

  @DISPID(3350)
  @PropGet
  boolean getKeepGeneralFormatDigitsWithEAsText();


  /**
   * <p>
   * Setter method for the COM property "KeepGeneralFormatDigitsWithEAsText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3350)
  @PropPut
  void setKeepGeneralFormatDigitsWithEAsText(
    boolean rhs);


  // Properties:
}
