package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ControlFormat extends Com4jObject {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addItem(text, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.String parameter.
   */

  @DISPID(851)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addItem(
    java.lang.String text);

  /**
   * @param text Mandatory java.lang.String parameter.
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(851)
  void addItem(
    java.lang.String text,
    @Optional java.lang.Object index);


  /**
   */

  @DISPID(853)
  void removeAllItems();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter count is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * removeItem(index, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   */

  @DISPID(852)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void removeItem(
    int index);

  /**
   * @param index Mandatory int parameter.
   * @param count Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(852)
  void removeItem(
    int index,
    @Optional java.lang.Object count);


  /**
   * <p>
   * Getter method for the COM property "DropDownLines"
   * </p>
   */

  @DISPID(848)
  @PropGet
  int getDropDownLines();


  /**
   * <p>
   * Setter method for the COM property "DropDownLines"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(848)
  @PropPut
  void setDropDownLines(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   */

  @DISPID(600)
  @PropGet
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(600)
  @PropPut
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LargeChange"
   * </p>
   */

  @DISPID(845)
  @PropGet
  int getLargeChange();


  /**
   * <p>
   * Setter method for the COM property "LargeChange"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(845)
  @PropPut
  void setLargeChange(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "LinkedCell"
   * </p>
   */

  @DISPID(1058)
  @PropGet
  java.lang.String getLinkedCell();


  /**
   * <p>
   * Setter method for the COM property "LinkedCell"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1058)
  @PropPut
  void setLinkedCell(
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * list(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(861)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object list();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(861)
  java.lang.Object list(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "ListCount"
   * </p>
   */

  @DISPID(849)
  @PropGet
  int getListCount();


  /**
   * <p>
   * Setter method for the COM property "ListCount"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(849)
  @PropPut
  void setListCount(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ListFillRange"
   * </p>
   */

  @DISPID(847)
  @PropGet
  java.lang.String getListFillRange();


  /**
   * <p>
   * Setter method for the COM property "ListFillRange"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(847)
  @PropPut
  void setListFillRange(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ListIndex"
   * </p>
   */

  @DISPID(850)
  @PropGet
  int getListIndex();


  /**
   * <p>
   * Setter method for the COM property "ListIndex"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(850)
  @PropPut
  void setListIndex(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "LockedText"
   * </p>
   */

  @DISPID(616)
  @PropGet
  boolean getLockedText();


  /**
   * <p>
   * Setter method for the COM property "LockedText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(616)
  @PropPut
  void setLockedText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Max"
   * </p>
   */

  @DISPID(842)
  @PropGet
  int getMax();


  /**
   * <p>
   * Setter method for the COM property "Max"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(842)
  @PropPut
  void setMax(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Min"
   * </p>
   */

  @DISPID(843)
  @PropGet
  int getMin();


  /**
   * <p>
   * Setter method for the COM property "Min"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(843)
  @PropPut
  void setMin(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MultiSelect"
   * </p>
   */

  @DISPID(32)
  @PropGet
  int getMultiSelect();


  /**
   * <p>
   * Setter method for the COM property "MultiSelect"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(32)
  @PropPut
  void setMultiSelect(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintObject"
   * </p>
   */

  @DISPID(618)
  @PropGet
  boolean getPrintObject();


  /**
   * <p>
   * Setter method for the COM property "PrintObject"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(618)
  @PropPut
  void setPrintObject(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SmallChange"
   * </p>
   */

  @DISPID(844)
  @PropGet
  int getSmallChange();


  /**
   * <p>
   * Setter method for the COM property "SmallChange"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(844)
  @PropPut
  void setSmallChange(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  int get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  void set_Default(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  int getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    int rhs);


  // Properties:
}
