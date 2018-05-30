package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002442A-0001-0000-C000-000000000046}")
public interface IParameter extends Com4jObject {
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
   * Getter method for the COM property "DataType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlParameterDataType
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlParameterDataType getDataType();


  /**
   * <p>
   * Setter method for the COM property "DataType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlParameterDataType parameter.
   */

  @VTID(11)
  void setDataType(
    com.exceljava.com4j.excel.XlParameterDataType rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlParameterType
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlParameterType getType();


  /**
   * <p>
   * Getter method for the COM property "PromptString"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getPromptString();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue();


  /**
   * <p>
   * Getter method for the COM property "SourceRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(15)
  com.exceljava.com4j.excel.Range getSourceRange();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(16)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(17)
  void setName(
    java.lang.String rhs);


  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlParameterType parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @VTID(18)
  void setParam(
    com.exceljava.com4j.excel.XlParameterType type,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * <p>
   * Getter method for the COM property "RefreshOnChange"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  boolean getRefreshOnChange();


  /**
   * <p>
   * Setter method for the COM property "RefreshOnChange"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(20)
  void setRefreshOnChange(
    boolean rhs);


  // Properties:
}
