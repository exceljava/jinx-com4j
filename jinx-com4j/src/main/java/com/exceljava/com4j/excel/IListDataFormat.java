package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002447D-0001-0000-C000-000000000046}")
public interface IListDataFormat extends Com4jObject {
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
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlListDataType
   */

  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.excel.XlListDataType get_Default();


  /**
   * <p>
   * Getter method for the COM property "Choices"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChoices();


  /**
   * <p>
   * Getter method for the COM property "DecimalPlaces"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getDecimalPlaces();


  /**
   * <p>
   * Getter method for the COM property "DefaultValue"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDefaultValue();


  /**
   * <p>
   * Getter method for the COM property "IsPercent"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getIsPercent();


  /**
   * <p>
   * Getter method for the COM property "lcid"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(15)
  int getLcid();


  /**
   * <p>
   * Getter method for the COM property "MaxCharacters"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getMaxCharacters();


  /**
   * <p>
   * Getter method for the COM property "MaxNumber"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(17)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getMaxNumber();


  /**
   * <p>
   * Getter method for the COM property "MinNumber"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getMinNumber();


  /**
   * <p>
   * Getter method for the COM property "Required"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  boolean getRequired();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlListDataType
   */

  @VTID(20)
  com.exceljava.com4j.excel.XlListDataType getType();


  /**
   * <p>
   * Getter method for the COM property "ReadOnly"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(21)
  boolean getReadOnly();


  /**
   * <p>
   * Getter method for the COM property "AllowFillIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getAllowFillIn();


  // Properties:
}
