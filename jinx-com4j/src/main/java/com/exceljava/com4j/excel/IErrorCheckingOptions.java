package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002445B-0001-0000-C000-000000000046}")
public interface IErrorCheckingOptions extends Com4jObject {
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
   * Getter method for the COM property "BackgroundChecking"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getBackgroundChecking();


  /**
   * <p>
   * Setter method for the COM property "BackgroundChecking"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setBackgroundChecking(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IndicatorColorIndex"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlColorIndex
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlColorIndex getIndicatorColorIndex();


  /**
   * <p>
   * Setter method for the COM property "IndicatorColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @VTID(13)
  void setIndicatorColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "EvaluateToError"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getEvaluateToError();


  /**
   * <p>
   * Setter method for the COM property "EvaluateToError"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setEvaluateToError(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextDate"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getTextDate();


  /**
   * <p>
   * Setter method for the COM property "TextDate"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setTextDate(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberAsText"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getNumberAsText();


  /**
   * <p>
   * Setter method for the COM property "NumberAsText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setNumberAsText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "InconsistentFormula"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getInconsistentFormula();


  /**
   * <p>
   * Setter method for the COM property "InconsistentFormula"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setInconsistentFormula(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "OmittedCells"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getOmittedCells();


  /**
   * <p>
   * Setter method for the COM property "OmittedCells"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setOmittedCells(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UnlockedFormulaCells"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getUnlockedFormulaCells();


  /**
   * <p>
   * Setter method for the COM property "UnlockedFormulaCells"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setUnlockedFormulaCells(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EmptyCellReferences"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getEmptyCellReferences();


  /**
   * <p>
   * Setter method for the COM property "EmptyCellReferences"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setEmptyCellReferences(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListDataValidation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getListDataValidation();


  /**
   * <p>
   * Setter method for the COM property "ListDataValidation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setListDataValidation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "InconsistentTableFormula"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getInconsistentTableFormula();


  /**
   * <p>
   * Setter method for the COM property "InconsistentTableFormula"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setInconsistentTableFormula(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MisleadingNumberFormats"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(32)
  boolean getMisleadingNumberFormats();


  /**
   * <p>
   * Setter method for the COM property "MisleadingNumberFormats"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(33)
  void setMisleadingNumberFormats(
    boolean rhs);


  // Properties:
}
