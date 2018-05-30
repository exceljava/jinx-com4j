package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208D4-0001-0000-C000-000000000046}")
public interface IAutoCorrect extends Com4jObject {
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
   * @param what Mandatory java.lang.String parameter.
   * @param replacement Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object addReplacement(
    java.lang.String what,
    java.lang.String replacement);


  /**
   * <p>
   * Getter method for the COM property "CapitalizeNamesOfDays"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getCapitalizeNamesOfDays();


  /**
   * <p>
   * Setter method for the COM property "CapitalizeNamesOfDays"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(12)
  void setCapitalizeNamesOfDays(
    boolean rhs);


  /**
   * @param what Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getReplacementList();

  /**
   * <p>
   * Getter method for the COM property "ReplacementList"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getReplacementList(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


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

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setReplacementList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "ReplacementList"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setReplacementList(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ReplaceText"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getReplaceText();


  /**
   * <p>
   * Setter method for the COM property "ReplaceText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setReplaceText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TwoInitialCapitals"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getTwoInitialCapitals();


  /**
   * <p>
   * Setter method for the COM property "TwoInitialCapitals"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setTwoInitialCapitals(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CorrectSentenceCap"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getCorrectSentenceCap();


  /**
   * <p>
   * Setter method for the COM property "CorrectSentenceCap"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setCorrectSentenceCap(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CorrectCapsLock"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getCorrectCapsLock();


  /**
   * <p>
   * Setter method for the COM property "CorrectCapsLock"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setCorrectCapsLock(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAutoCorrectOptions"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getDisplayAutoCorrectOptions();


  /**
   * <p>
   * Setter method for the COM property "DisplayAutoCorrectOptions"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setDisplayAutoCorrectOptions(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoExpandListRange"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getAutoExpandListRange();


  /**
   * <p>
   * Setter method for the COM property "AutoExpandListRange"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setAutoExpandListRange(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFillFormulasInLists"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getAutoFillFormulasInLists();


  /**
   * <p>
   * Setter method for the COM property "AutoFillFormulasInLists"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setAutoFillFormulasInLists(
    boolean rhs);


  // Properties:
}
