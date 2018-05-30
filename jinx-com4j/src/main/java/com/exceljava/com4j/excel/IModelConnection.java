package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244D1-0001-0000-C000-000000000046}")
public interface IModelConnection extends Com4jObject {
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
   * Getter method for the COM property "CommandText"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCommandText();


  /**
   * <p>
   * Setter method for the COM property "CommandText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(11)
  void setCommandText(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCmdType
   */

  @VTID(12)
  com.exceljava.com4j.excel.XlCmdType getCommandType();


  /**
   * <p>
   * Setter method for the COM property "CommandType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCmdType parameter.
   */

  @VTID(13)
  void setCommandType(
    com.exceljava.com4j.excel.XlCmdType rhs);


  /**
   * <p>
   * Getter method for the COM property "ADOConnection"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getADOConnection();


  /**
   * <p>
   * Getter method for the COM property "CalculatedMembers"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMembers
   */

  @VTID(15)
  com.exceljava.com4j.excel.CalculatedMembers getCalculatedMembers();


  // Properties:
}
