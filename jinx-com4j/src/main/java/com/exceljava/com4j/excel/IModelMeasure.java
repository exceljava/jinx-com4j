package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244ED-0001-0000-C000-000000000046}")
public interface IModelMeasure extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(11)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "AssociatedTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTable
   */

  @VTID(12)
  com.exceljava.com4j.excel.ModelTable getAssociatedTable();


  /**
   * <p>
   * Setter method for the COM property "AssociatedTable"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.ModelTable parameter.
   */

  @VTID(13)
  void setAssociatedTable(
    com.exceljava.com4j.excel.ModelTable rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(15)
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FormatInformation"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getFormatInformation();


  /**
   * <p>
   * Setter method for the COM property "FormatInformation"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(17)
  void setFormatInformation(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(18)
  java.lang.String getDescription();


  /**
   * <p>
   * Setter method for the COM property "Description"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(19)
  void setDescription(
    java.lang.String rhs);


  /**
   */

  @VTID(20)
  void delete();


  // Properties:
}
