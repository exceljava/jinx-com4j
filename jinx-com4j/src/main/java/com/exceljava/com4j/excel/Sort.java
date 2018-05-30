package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Sort extends Com4jObject {
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
   * Getter method for the COM property "Rng"
   * </p>
   */

  @DISPID(2748)
  @PropGet
  com.exceljava.com4j.excel.Range getRng();


  /**
   * <p>
   * Getter method for the COM property "Header"
   * </p>
   */

  @DISPID(895)
  @PropGet
  com.exceljava.com4j.excel.XlYesNoGuess getHeader();


  /**
   * <p>
   * Setter method for the COM property "Header"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlYesNoGuess parameter.
   */

  @DISPID(895)
  @PropPut
  void setHeader(
    com.exceljava.com4j.excel.XlYesNoGuess rhs);


  /**
   * <p>
   * Getter method for the COM property "MatchCase"
   * </p>
   */

  @DISPID(426)
  @PropGet
  boolean getMatchCase();


  /**
   * <p>
   * Setter method for the COM property "MatchCase"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(426)
  @PropPut
  void setMatchCase(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  com.exceljava.com4j.excel.XlSortOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortOrientation parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    com.exceljava.com4j.excel.XlSortOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "SortMethod"
   * </p>
   */

  @DISPID(897)
  @PropGet
  com.exceljava.com4j.excel.XlSortMethod getSortMethod();


  /**
   * <p>
   * Setter method for the COM property "SortMethod"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSortMethod parameter.
   */

  @DISPID(897)
  @PropPut
  void setSortMethod(
    com.exceljava.com4j.excel.XlSortMethod rhs);


  /**
   * <p>
   * Getter method for the COM property "SortFields"
   * </p>
   */

  @DISPID(2749)
  @PropGet
  com.exceljava.com4j.excel.SortFields getSortFields();


  /**
   * @param rng Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(2750)
  void setRange(
    com.exceljava.com4j.excel.Range rng);


  /**
   */

  @DISPID(1675)
  void apply();


  // Properties:
}
