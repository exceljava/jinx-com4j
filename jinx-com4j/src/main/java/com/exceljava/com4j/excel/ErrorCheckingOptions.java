package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ErrorCheckingOptions extends Com4jObject {
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
   * Getter method for the COM property "BackgroundChecking"
   * </p>
   */

  @DISPID(2201)
  @PropGet
  boolean getBackgroundChecking();


  /**
   * <p>
   * Setter method for the COM property "BackgroundChecking"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2201)
  @PropPut
  void setBackgroundChecking(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IndicatorColorIndex"
   * </p>
   */

  @DISPID(2202)
  @PropGet
  com.exceljava.com4j.excel.XlColorIndex getIndicatorColorIndex();


  /**
   * <p>
   * Setter method for the COM property "IndicatorColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @DISPID(2202)
  @PropPut
  void setIndicatorColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "EvaluateToError"
   * </p>
   */

  @DISPID(2203)
  @PropGet
  boolean getEvaluateToError();


  /**
   * <p>
   * Setter method for the COM property "EvaluateToError"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2203)
  @PropPut
  void setEvaluateToError(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextDate"
   * </p>
   */

  @DISPID(2204)
  @PropGet
  boolean getTextDate();


  /**
   * <p>
   * Setter method for the COM property "TextDate"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2204)
  @PropPut
  void setTextDate(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberAsText"
   * </p>
   */

  @DISPID(2205)
  @PropGet
  boolean getNumberAsText();


  /**
   * <p>
   * Setter method for the COM property "NumberAsText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2205)
  @PropPut
  void setNumberAsText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "InconsistentFormula"
   * </p>
   */

  @DISPID(2206)
  @PropGet
  boolean getInconsistentFormula();


  /**
   * <p>
   * Setter method for the COM property "InconsistentFormula"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2206)
  @PropPut
  void setInconsistentFormula(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "OmittedCells"
   * </p>
   */

  @DISPID(2207)
  @PropGet
  boolean getOmittedCells();


  /**
   * <p>
   * Setter method for the COM property "OmittedCells"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2207)
  @PropPut
  void setOmittedCells(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UnlockedFormulaCells"
   * </p>
   */

  @DISPID(2208)
  @PropGet
  boolean getUnlockedFormulaCells();


  /**
   * <p>
   * Setter method for the COM property "UnlockedFormulaCells"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2208)
  @PropPut
  void setUnlockedFormulaCells(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EmptyCellReferences"
   * </p>
   */

  @DISPID(2209)
  @PropGet
  boolean getEmptyCellReferences();


  /**
   * <p>
   * Setter method for the COM property "EmptyCellReferences"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2209)
  @PropPut
  void setEmptyCellReferences(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListDataValidation"
   * </p>
   */

  @DISPID(2296)
  @PropGet
  boolean getListDataValidation();


  /**
   * <p>
   * Setter method for the COM property "ListDataValidation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2296)
  @PropPut
  void setListDataValidation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "InconsistentTableFormula"
   * </p>
   */

  @DISPID(2675)
  @PropGet
  boolean getInconsistentTableFormula();


  /**
   * <p>
   * Setter method for the COM property "InconsistentTableFormula"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2675)
  @PropPut
  void setInconsistentTableFormula(
    boolean rhs);


  // Properties:
}
