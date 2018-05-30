package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ListDataFormat extends Com4jObject {
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
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.XlListDataType get_Default();


  /**
   * <p>
   * Getter method for the COM property "Choices"
   * </p>
   */

  @DISPID(2348)
  @PropGet
  java.lang.Object getChoices();


  /**
   * <p>
   * Getter method for the COM property "DecimalPlaces"
   * </p>
   */

  @DISPID(2349)
  @PropGet
  int getDecimalPlaces();


  /**
   * <p>
   * Getter method for the COM property "DefaultValue"
   * </p>
   */

  @DISPID(2350)
  @PropGet
  java.lang.Object getDefaultValue();


  /**
   * <p>
   * Getter method for the COM property "IsPercent"
   * </p>
   */

  @DISPID(2351)
  @PropGet
  boolean getIsPercent();


  /**
   * <p>
   * Getter method for the COM property "lcid"
   * </p>
   */

  @DISPID(2352)
  @PropGet
  int getLcid();


  /**
   * <p>
   * Getter method for the COM property "MaxCharacters"
   * </p>
   */

  @DISPID(2353)
  @PropGet
  int getMaxCharacters();


  /**
   * <p>
   * Getter method for the COM property "MaxNumber"
   * </p>
   */

  @DISPID(2354)
  @PropGet
  java.lang.Object getMaxNumber();


  /**
   * <p>
   * Getter method for the COM property "MinNumber"
   * </p>
   */

  @DISPID(2355)
  @PropGet
  java.lang.Object getMinNumber();


  /**
   * <p>
   * Getter method for the COM property "Required"
   * </p>
   */

  @DISPID(2356)
  @PropGet
  boolean getRequired();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.excel.XlListDataType getType();


  /**
   * <p>
   * Getter method for the COM property "ReadOnly"
   * </p>
   */

  @DISPID(296)
  @PropGet
  boolean getReadOnly();


  /**
   * <p>
   * Getter method for the COM property "AllowFillIn"
   * </p>
   */

  @DISPID(2357)
  @PropGet
  boolean getAllowFillIn();


  // Properties:
}
