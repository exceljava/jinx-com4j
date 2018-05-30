package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelMeasure extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "AssociatedTable"
   * </p>
   */

  @DISPID(3224)
  @PropGet
  com.exceljava.com4j.excel.ModelTable getAssociatedTable();


  /**
   * <p>
   * Setter method for the COM property "AssociatedTable"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.ModelTable parameter.
   */

  @DISPID(3224)
  @PropPut
  void setAssociatedTable(
    com.exceljava.com4j.excel.ModelTable rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   */

  @DISPID(261)
  @PropGet
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(261)
  @PropPut
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FormatInformation"
   * </p>
   */

  @DISPID(3225)
  @PropGet
  java.lang.Object getFormatInformation();


  /**
   * <p>
   * Setter method for the COM property "FormatInformation"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3225)
  @PropPut
  void setFormatInformation(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   */

  @DISPID(218)
  @PropGet
  java.lang.String getDescription();


  /**
   * <p>
   * Setter method for the COM property "Description"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(218)
  @PropPut
  void setDescription(
    java.lang.String rhs);


  /**
   */

  @DISPID(117)
  void delete();


  // Properties:
}
