package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Name extends Com4jObject {
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
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   */

  @DISPID(486)
  @PropGet
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Category"
   * </p>
   */

  @DISPID(934)
  @PropGet
  java.lang.String getCategory();


  /**
   * <p>
   * Setter method for the COM property "Category"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(934)
  @PropPut
  void setCategory(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "CategoryLocal"
   * </p>
   */

  @DISPID(935)
  @PropGet
  java.lang.String getCategoryLocal();


  /**
   * <p>
   * Setter method for the COM property "CategoryLocal"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(935)
  @PropPut
  void setCategoryLocal(
    java.lang.String rhs);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "MacroType"
   * </p>
   */

  @DISPID(936)
  @PropGet
  com.exceljava.com4j.excel.XlXLMMacroType getMacroType();


  /**
   * <p>
   * Setter method for the COM property "MacroType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlXLMMacroType parameter.
   */

  @DISPID(936)
  @PropPut
  void setMacroType(
    com.exceljava.com4j.excel.XlXLMMacroType rhs);


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
   * Getter method for the COM property "RefersTo"
   * </p>
   */

  @DISPID(938)
  @PropGet
  java.lang.Object getRefersTo();


  /**
   * <p>
   * Setter method for the COM property "RefersTo"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(938)
  @PropPut
  void setRefersTo(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ShortcutKey"
   * </p>
   */

  @DISPID(597)
  @PropGet
  java.lang.String getShortcutKey();


  /**
   * <p>
   * Setter method for the COM property "ShortcutKey"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(597)
  @PropPut
  void setShortcutKey(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.String getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "NameLocal"
   * </p>
   */

  @DISPID(937)
  @PropGet
  java.lang.String getNameLocal();


  /**
   * <p>
   * Setter method for the COM property "NameLocal"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(937)
  @PropPut
  void setNameLocal(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RefersToLocal"
   * </p>
   */

  @DISPID(939)
  @PropGet
  java.lang.Object getRefersToLocal();


  /**
   * <p>
   * Setter method for the COM property "RefersToLocal"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(939)
  @PropPut
  void setRefersToLocal(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "RefersToR1C1"
   * </p>
   */

  @DISPID(940)
  @PropGet
  java.lang.Object getRefersToR1C1();


  /**
   * <p>
   * Setter method for the COM property "RefersToR1C1"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(940)
  @PropPut
  void setRefersToR1C1(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "RefersToR1C1Local"
   * </p>
   */

  @DISPID(941)
  @PropGet
  java.lang.Object getRefersToR1C1Local();


  /**
   * <p>
   * Setter method for the COM property "RefersToR1C1Local"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(941)
  @PropPut
  void setRefersToR1C1Local(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "RefersToRange"
   * </p>
   */

  @DISPID(1160)
  @PropGet
  com.exceljava.com4j.excel.Range getRefersToRange();


  /**
   * <p>
   * Getter method for the COM property "Comment"
   * </p>
   */

  @DISPID(910)
  @PropGet
  java.lang.String getComment();


  /**
   * <p>
   * Setter method for the COM property "Comment"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(910)
  @PropPut
  void setComment(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "WorkbookParameter"
   * </p>
   */

  @DISPID(2607)
  @PropGet
  boolean getWorkbookParameter();


  /**
   * <p>
   * Setter method for the COM property "WorkbookParameter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2607)
  @PropPut
  void setWorkbookParameter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ValidWorkbookParameter"
   * </p>
   */

  @DISPID(2608)
  @PropGet
  boolean getValidWorkbookParameter();


  // Properties:
}
