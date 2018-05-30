package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020874-0001-0000-C000-000000000046}")
public interface IPivotField extends Com4jObject {
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
   * Getter method for the COM property "Calculation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotFieldCalculation
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlPivotFieldCalculation getCalculation();


  /**
   * <p>
   * Setter method for the COM property "Calculation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotFieldCalculation parameter.
   */

  @VTID(11)
  void setCalculation(
    com.exceljava.com4j.excel.XlPivotFieldCalculation rhs);


  /**
   * <p>
   * Getter method for the COM property "ChildField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(12)
  com.exceljava.com4j.excel.PivotField getChildField();


  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChildItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getChildItems();

  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChildItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "CurrentPage"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCurrentPage();


  /**
   * <p>
   * Setter method for the COM property "CurrentPage"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setCurrentPage(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DataRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(16)
  com.exceljava.com4j.excel.Range getDataRange();


  /**
   * <p>
   * Getter method for the COM property "DataType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotFieldDataType
   */

  @VTID(17)
  com.exceljava.com4j.excel.XlPivotFieldDataType getDataType();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(18)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(19)
  @DefaultMethod
  void set_Default(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Function"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlConsolidationFunction
   */

  @VTID(20)
  com.exceljava.com4j.excel.XlConsolidationFunction getFunction();


  /**
   * <p>
   * Setter method for the COM property "Function"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   */

  @VTID(21)
  void setFunction(
    com.exceljava.com4j.excel.XlConsolidationFunction rhs);


  /**
   * <p>
   * Getter method for the COM property "GroupLevel"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getGroupLevel();


  /**
   * <p>
   * Getter method for the COM property "HiddenItems"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHiddenItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getHiddenItems();

  /**
   * <p>
   * Getter method for the COM property "HiddenItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHiddenItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "LabelRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(24)
  com.exceljava.com4j.excel.Range getLabelRange();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(25)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(26)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(28)
  void setNumberFormat(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotFieldOrientation
   */

  @VTID(29)
  com.exceljava.com4j.excel.XlPivotFieldOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotFieldOrientation parameter.
   */

  @VTID(30)
  void setOrientation(
    com.exceljava.com4j.excel.XlPivotFieldOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAllItems"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getShowAllItems();


  /**
   * <p>
   * Setter method for the COM property "ShowAllItems"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setShowAllItems(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ParentField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(33)
  com.exceljava.com4j.excel.PivotField getParentField();


  /**
   * <p>
   * Getter method for the COM property "ParentItems"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getParentItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getParentItems();

  /**
   * <p>
   * Getter method for the COM property "ParentItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getParentItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object pivotItems();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(35)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object pivotItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(36)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(37)
  void setPosition(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(38)
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "Subtotals"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSubtotals(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getSubtotals();

  /**
   * <p>
   * Getter method for the COM property "Subtotals"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(39)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSubtotals(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "Subtotals"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSubtotals(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(40)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSubtotals(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Subtotals"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(40)
  void setSubtotals(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BaseField"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(41)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getBaseField();


  /**
   * <p>
   * Setter method for the COM property "BaseField"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(42)
  void setBaseField(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BaseItem"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(43)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getBaseItem();


  /**
   * <p>
   * Setter method for the COM property "BaseItem"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(44)
  void setBaseItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TotalLevels"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(45)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTotalLevels();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(46)
  java.lang.String getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(47)
  void setValue(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleItems"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getVisibleItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(48)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getVisibleItems();

  /**
   * <p>
   * Getter method for the COM property "VisibleItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(48)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVisibleItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedItems
   */

  @VTID(49)
  com.exceljava.com4j.excel.CalculatedItems calculatedItems();


  /**
   */

  @VTID(50)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "DragToColumn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(51)
  boolean getDragToColumn();


  /**
   * <p>
   * Setter method for the COM property "DragToColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(52)
  void setDragToColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToHide"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(53)
  boolean getDragToHide();


  /**
   * <p>
   * Setter method for the COM property "DragToHide"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(54)
  void setDragToHide(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToPage"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(55)
  boolean getDragToPage();


  /**
   * <p>
   * Setter method for the COM property "DragToPage"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(56)
  void setDragToPage(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToRow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(57)
  boolean getDragToRow();


  /**
   * <p>
   * Setter method for the COM property "DragToRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(58)
  void setDragToRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(59)
  boolean getDragToData();


  /**
   * <p>
   * Setter method for the COM property "DragToData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(60)
  void setDragToData(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(61)
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(62)
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "IsCalculated"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(63)
  boolean getIsCalculated();


  /**
   * <p>
   * Getter method for the COM property "MemoryUsed"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(64)
  int getMemoryUsed();


  /**
   * <p>
   * Getter method for the COM property "ServerBased"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(65)
  boolean getServerBased();


  /**
   * <p>
   * Setter method for the COM property "ServerBased"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(66)
  void setServerBased(
    boolean rhs);


  /**
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   */

  @VTID(67)
  void _AutoSort(
    int order,
    java.lang.String field);


  /**
   * @param type Mandatory int parameter.
   * @param range Mandatory int parameter.
   * @param count Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   */

  @VTID(68)
  void autoShow(
    int type,
    int range,
    int count,
    java.lang.String field);


  /**
   * <p>
   * Getter method for the COM property "AutoSortOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(69)
  int getAutoSortOrder();


  /**
   * <p>
   * Getter method for the COM property "AutoSortField"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(70)
  java.lang.String getAutoSortField();


  /**
   * <p>
   * Getter method for the COM property "AutoShowType"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(71)
  int getAutoShowType();


  /**
   * <p>
   * Getter method for the COM property "AutoShowRange"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(72)
  int getAutoShowRange();


  /**
   * <p>
   * Getter method for the COM property "AutoShowCount"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(73)
  int getAutoShowCount();


  /**
   * <p>
   * Getter method for the COM property "AutoShowField"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(74)
  java.lang.String getAutoShowField();


  /**
   * <p>
   * Getter method for the COM property "LayoutBlankLine"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(75)
  boolean getLayoutBlankLine();


  /**
   * <p>
   * Setter method for the COM property "LayoutBlankLine"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(76)
  void setLayoutBlankLine(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutSubtotalLocation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSubtototalLocationType
   */

  @VTID(77)
  com.exceljava.com4j.excel.XlSubtototalLocationType getLayoutSubtotalLocation();


  /**
   * <p>
   * Setter method for the COM property "LayoutSubtotalLocation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSubtototalLocationType parameter.
   */

  @VTID(78)
  void setLayoutSubtotalLocation(
    com.exceljava.com4j.excel.XlSubtototalLocationType rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutPageBreak"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(79)
  boolean getLayoutPageBreak();


  /**
   * <p>
   * Setter method for the COM property "LayoutPageBreak"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(80)
  void setLayoutPageBreak(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutForm"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlLayoutFormType
   */

  @VTID(81)
  com.exceljava.com4j.excel.XlLayoutFormType getLayoutForm();


  /**
   * <p>
   * Setter method for the COM property "LayoutForm"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlLayoutFormType parameter.
   */

  @VTID(82)
  void setLayoutForm(
    com.exceljava.com4j.excel.XlLayoutFormType rhs);


  /**
   * <p>
   * Getter method for the COM property "SubtotalName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(83)
  java.lang.String getSubtotalName();


  /**
   * <p>
   * Setter method for the COM property "SubtotalName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(84)
  void setSubtotalName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(85)
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(86)
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DrilledDown"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(87)
  boolean getDrilledDown();


  /**
   * <p>
   * Setter method for the COM property "DrilledDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(88)
  void setDrilledDown(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CubeField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CubeField
   */

  @VTID(89)
  com.exceljava.com4j.excel.CubeField getCubeField();


  /**
   * <p>
   * Getter method for the COM property "CurrentPageName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(90)
  java.lang.String getCurrentPageName();


  /**
   * <p>
   * Setter method for the COM property "CurrentPageName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(91)
  void setCurrentPageName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "StandardFormula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(92)
  java.lang.String getStandardFormula();


  /**
   * <p>
   * Setter method for the COM property "StandardFormula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(93)
  void setStandardFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "HiddenItemsList"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(94)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHiddenItemsList();


  /**
   * <p>
   * Setter method for the COM property "HiddenItemsList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(95)
  void setHiddenItemsList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DatabaseSort"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(96)
  boolean getDatabaseSort();


  /**
   * <p>
   * Setter method for the COM property "DatabaseSort"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(97)
  void setDatabaseSort(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IsMemberProperty"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(98)
  boolean getIsMemberProperty();


  /**
   * <p>
   * Getter method for the COM property "PropertyParentField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(99)
  com.exceljava.com4j.excel.PivotField getPropertyParentField();


  /**
   * <p>
   * Getter method for the COM property "PropertyOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(100)
  int getPropertyOrder();


  /**
   * <p>
   * Setter method for the COM property "PropertyOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(101)
  void setPropertyOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableItemSelection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(102)
  boolean getEnableItemSelection();


  /**
   * <p>
   * Setter method for the COM property "EnableItemSelection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(103)
  void setEnableItemSelection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CurrentPageList"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCurrentPageList();


  /**
   * <p>
   * Setter method for the COM property "CurrentPageList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(105)
  void setCurrentPageList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter clearList is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPageItem(item, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param item Mandatory java.lang.String parameter.
   */

  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addPageItem(
    java.lang.String item);

  /**
   * @param item Mandatory java.lang.String parameter.
   * @param clearList Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(106)
  void addPageItem(
    java.lang.String item,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object clearList);


  /**
   * <p>
   * Getter method for the COM property "Hidden"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(107)
  boolean getHidden();


  /**
   * <p>
   * Setter method for the COM property "Hidden"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(108)
  void setHidden(
    boolean rhs);


  /**
   * @param field Mandatory java.lang.String parameter.
   */

  @VTID(109)
  void drillTo(
    java.lang.String field);


  /**
   * <p>
   * Getter method for the COM property "UseMemberPropertyAsCaption"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(110)
  boolean getUseMemberPropertyAsCaption();


  /**
   * <p>
   * Setter method for the COM property "UseMemberPropertyAsCaption"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(111)
  void setUseMemberPropertyAsCaption(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MemberPropertyCaption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(112)
  java.lang.String getMemberPropertyCaption();


  /**
   * <p>
   * Setter method for the COM property "MemberPropertyCaption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(113)
  void setMemberPropertyCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAsTooltip"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(114)
  boolean getDisplayAsTooltip();


  /**
   * <p>
   * Setter method for the COM property "DisplayAsTooltip"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(115)
  void setDisplayAsTooltip(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayInReport"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(116)
  boolean getDisplayInReport();


  /**
   * <p>
   * Setter method for the COM property "DisplayInReport"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(117)
  void setDisplayInReport(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAsCaption"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(118)
  boolean getDisplayAsCaption();


  /**
   * <p>
   * Getter method for the COM property "LayoutCompactRow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(119)
  boolean getLayoutCompactRow();


  /**
   * <p>
   * Setter method for the COM property "LayoutCompactRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(120)
  void setLayoutCompactRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeNewItemsInFilter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(121)
  boolean getIncludeNewItemsInFilter();


  /**
   * <p>
   * Setter method for the COM property "IncludeNewItemsInFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(122)
  void setIncludeNewItemsInFilter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleItemsList"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(123)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVisibleItemsList();


  /**
   * <p>
   * Setter method for the COM property "VisibleItemsList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(124)
  void setVisibleItemsList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PivotFilters"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilters
   */

  @VTID(125)
  com.exceljava.com4j.excel.PivotFilters getPivotFilters();


  /**
   * <p>
   * Getter method for the COM property "AutoSortPivotLine"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotLine
   */

  @VTID(126)
  com.exceljava.com4j.excel.PivotLine getAutoSortPivotLine();


  /**
   * <p>
   * Getter method for the COM property "AutoSortCustomSubtotal"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(127)
  int getAutoSortCustomSubtotal();


  /**
   * <p>
   * Getter method for the COM property "ShowingInAxis"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(128)
  boolean getShowingInAxis();


  /**
   * <p>
   * Getter method for the COM property "EnableMultiplePageItems"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(129)
  boolean getEnableMultiplePageItems();


  /**
   * <p>
   * Setter method for the COM property "EnableMultiplePageItems"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(130)
  void setEnableMultiplePageItems(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AllItemsVisible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(131)
  boolean getAllItemsVisible();


  /**
   */

  @VTID(132)
  void clearManualFilter();


  /**
   */

  @VTID(133)
  void clearAllFilters();


  /**
   */

  @VTID(134)
  void clearValueFilters();


  /**
   */

  @VTID(135)
  void clearLabelFilters();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pivotLine is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter customSubtotal is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoSort(order, field, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   */

  @VTID(136)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void autoSort(
    int order,
    java.lang.String field);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customSubtotal is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoSort(order, field, pivotLine, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   * @param pivotLine Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(136)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void autoSort(
    int order,
    java.lang.String field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pivotLine);

  /**
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   * @param pivotLine Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customSubtotal Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(136)
  void autoSort(
    int order,
    java.lang.String field,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pivotLine,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customSubtotal);


  /**
   * <p>
   * Getter method for the COM property "SourceCaption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(137)
  java.lang.String getSourceCaption();


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(138)
  boolean getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(139)
  void setShowDetail(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RepeatLabels"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(140)
  boolean getRepeatLabels();


  /**
   * <p>
   * Setter method for the COM property "RepeatLabels"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(141)
  void setRepeatLabels(
    boolean rhs);


  /**
   */

  @VTID(142)
  void autoGroup();


  // Properties:
}
