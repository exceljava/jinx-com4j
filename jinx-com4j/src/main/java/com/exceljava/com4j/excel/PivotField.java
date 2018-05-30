package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotField extends Com4jObject {
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
   * Getter method for the COM property "Calculation"
   * </p>
   */

  @DISPID(316)
  @PropGet
  com.exceljava.com4j.excel.XlPivotFieldCalculation getCalculation();


  /**
   * <p>
   * Setter method for the COM property "Calculation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotFieldCalculation parameter.
   */

  @DISPID(316)
  @PropPut
  void setCalculation(
    com.exceljava.com4j.excel.XlPivotFieldCalculation rhs);


  /**
   * <p>
   * Getter method for the COM property "ChildField"
   * </p>
   */

  @DISPID(736)
  @PropGet
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
   */

  @DISPID(730)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getChildItems();

  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(730)
  @PropGet
  java.lang.Object getChildItems(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "CurrentPage"
   * </p>
   */

  @DISPID(738)
  @PropGet
  java.lang.Object getCurrentPage();


  /**
   * <p>
   * Setter method for the COM property "CurrentPage"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(738)
  @PropPut
  void setCurrentPage(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DataRange"
   * </p>
   */

  @DISPID(720)
  @PropGet
  com.exceljava.com4j.excel.Range getDataRange();


  /**
   * <p>
   * Getter method for the COM property "DataType"
   * </p>
   */

  @DISPID(722)
  @PropGet
  com.exceljava.com4j.excel.XlPivotFieldDataType getDataType();


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
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  void set_Default(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Function"
   * </p>
   */

  @DISPID(899)
  @PropGet
  com.exceljava.com4j.excel.XlConsolidationFunction getFunction();


  /**
   * <p>
   * Setter method for the COM property "Function"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlConsolidationFunction parameter.
   */

  @DISPID(899)
  @PropPut
  void setFunction(
    com.exceljava.com4j.excel.XlConsolidationFunction rhs);


  /**
   * <p>
   * Getter method for the COM property "GroupLevel"
   * </p>
   */

  @DISPID(723)
  @PropGet
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
   */

  @DISPID(728)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getHiddenItems();

  /**
   * <p>
   * Getter method for the COM property "HiddenItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(728)
  @PropGet
  java.lang.Object getHiddenItems(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "LabelRange"
   * </p>
   */

  @DISPID(719)
  @PropGet
  com.exceljava.com4j.excel.Range getLabelRange();


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
   * Getter method for the COM property "NumberFormat"
   * </p>
   */

  @DISPID(193)
  @PropGet
  java.lang.String getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(193)
  @PropPut
  void setNumberFormat(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  com.exceljava.com4j.excel.XlPivotFieldOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotFieldOrientation parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    com.exceljava.com4j.excel.XlPivotFieldOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAllItems"
   * </p>
   */

  @DISPID(452)
  @PropGet
  boolean getShowAllItems();


  /**
   * <p>
   * Setter method for the COM property "ShowAllItems"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(452)
  @PropPut
  void setShowAllItems(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ParentField"
   * </p>
   */

  @DISPID(732)
  @PropGet
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
   */

  @DISPID(729)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getParentItems();

  /**
   * <p>
   * Getter method for the COM property "ParentItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(729)
  @PropGet
  java.lang.Object getParentItems(
    @Optional java.lang.Object index);


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
   */

  @DISPID(737)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object pivotItems();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(737)
  java.lang.Object pivotItems(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  java.lang.Object getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(133)
  @PropPut
  void setPosition(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   */

  @DISPID(721)
  @PropGet
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
   */

  @DISPID(733)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getSubtotals();

  /**
   * <p>
   * Getter method for the COM property "Subtotals"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(733)
  @PropGet
  java.lang.Object getSubtotals(
    @Optional java.lang.Object index);


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

  @DISPID(733)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSubtotals(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Subtotals"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(733)
  @PropPut
  void setSubtotals(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BaseField"
   * </p>
   */

  @DISPID(734)
  @PropGet
  java.lang.Object getBaseField();


  /**
   * <p>
   * Setter method for the COM property "BaseField"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(734)
  @PropPut
  void setBaseField(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "BaseItem"
   * </p>
   */

  @DISPID(735)
  @PropGet
  java.lang.Object getBaseItem();


  /**
   * <p>
   * Setter method for the COM property "BaseItem"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(735)
  @PropPut
  void setBaseItem(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TotalLevels"
   * </p>
   */

  @DISPID(724)
  @PropGet
  java.lang.Object getTotalLevels();


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
   */

  @DISPID(727)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getVisibleItems();

  /**
   * <p>
   * Getter method for the COM property "VisibleItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(727)
  @PropGet
  java.lang.Object getVisibleItems(
    @Optional java.lang.Object index);


  /**
   */

  @DISPID(1507)
  com.exceljava.com4j.excel.CalculatedItems calculatedItems();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "DragToColumn"
   * </p>
   */

  @DISPID(1508)
  @PropGet
  boolean getDragToColumn();


  /**
   * <p>
   * Setter method for the COM property "DragToColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1508)
  @PropPut
  void setDragToColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToHide"
   * </p>
   */

  @DISPID(1509)
  @PropGet
  boolean getDragToHide();


  /**
   * <p>
   * Setter method for the COM property "DragToHide"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1509)
  @PropPut
  void setDragToHide(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToPage"
   * </p>
   */

  @DISPID(1510)
  @PropGet
  boolean getDragToPage();


  /**
   * <p>
   * Setter method for the COM property "DragToPage"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1510)
  @PropPut
  void setDragToPage(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToRow"
   * </p>
   */

  @DISPID(1511)
  @PropGet
  boolean getDragToRow();


  /**
   * <p>
   * Setter method for the COM property "DragToRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1511)
  @PropPut
  void setDragToRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DragToData"
   * </p>
   */

  @DISPID(1844)
  @PropGet
  boolean getDragToData();


  /**
   * <p>
   * Setter method for the COM property "DragToData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1844)
  @PropPut
  void setDragToData(
    boolean rhs);


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
   * Getter method for the COM property "IsCalculated"
   * </p>
   */

  @DISPID(1512)
  @PropGet
  boolean getIsCalculated();


  /**
   * <p>
   * Getter method for the COM property "MemoryUsed"
   * </p>
   */

  @DISPID(372)
  @PropGet
  int getMemoryUsed();


  /**
   * <p>
   * Getter method for the COM property "ServerBased"
   * </p>
   */

  @DISPID(1513)
  @PropGet
  boolean getServerBased();


  /**
   * <p>
   * Setter method for the COM property "ServerBased"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1513)
  @PropPut
  void setServerBased(
    boolean rhs);


  /**
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   */

  @DISPID(2579)
  void _AutoSort(
    int order,
    java.lang.String field);


  /**
   * @param type Mandatory int parameter.
   * @param range Mandatory int parameter.
   * @param count Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   */

  @DISPID(1515)
  void autoShow(
    int type,
    int range,
    int count,
    java.lang.String field);


  /**
   * <p>
   * Getter method for the COM property "AutoSortOrder"
   * </p>
   */

  @DISPID(1516)
  @PropGet
  int getAutoSortOrder();


  /**
   * <p>
   * Getter method for the COM property "AutoSortField"
   * </p>
   */

  @DISPID(1517)
  @PropGet
  java.lang.String getAutoSortField();


  /**
   * <p>
   * Getter method for the COM property "AutoShowType"
   * </p>
   */

  @DISPID(1518)
  @PropGet
  int getAutoShowType();


  /**
   * <p>
   * Getter method for the COM property "AutoShowRange"
   * </p>
   */

  @DISPID(1519)
  @PropGet
  int getAutoShowRange();


  /**
   * <p>
   * Getter method for the COM property "AutoShowCount"
   * </p>
   */

  @DISPID(1520)
  @PropGet
  int getAutoShowCount();


  /**
   * <p>
   * Getter method for the COM property "AutoShowField"
   * </p>
   */

  @DISPID(1521)
  @PropGet
  java.lang.String getAutoShowField();


  /**
   * <p>
   * Getter method for the COM property "LayoutBlankLine"
   * </p>
   */

  @DISPID(1845)
  @PropGet
  boolean getLayoutBlankLine();


  /**
   * <p>
   * Setter method for the COM property "LayoutBlankLine"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1845)
  @PropPut
  void setLayoutBlankLine(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutSubtotalLocation"
   * </p>
   */

  @DISPID(1846)
  @PropGet
  com.exceljava.com4j.excel.XlSubtototalLocationType getLayoutSubtotalLocation();


  /**
   * <p>
   * Setter method for the COM property "LayoutSubtotalLocation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSubtototalLocationType parameter.
   */

  @DISPID(1846)
  @PropPut
  void setLayoutSubtotalLocation(
    com.exceljava.com4j.excel.XlSubtototalLocationType rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutPageBreak"
   * </p>
   */

  @DISPID(1847)
  @PropGet
  boolean getLayoutPageBreak();


  /**
   * <p>
   * Setter method for the COM property "LayoutPageBreak"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1847)
  @PropPut
  void setLayoutPageBreak(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LayoutForm"
   * </p>
   */

  @DISPID(1848)
  @PropGet
  com.exceljava.com4j.excel.XlLayoutFormType getLayoutForm();


  /**
   * <p>
   * Setter method for the COM property "LayoutForm"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlLayoutFormType parameter.
   */

  @DISPID(1848)
  @PropPut
  void setLayoutForm(
    com.exceljava.com4j.excel.XlLayoutFormType rhs);


  /**
   * <p>
   * Getter method for the COM property "SubtotalName"
   * </p>
   */

  @DISPID(1849)
  @PropGet
  java.lang.String getSubtotalName();


  /**
   * <p>
   * Setter method for the COM property "SubtotalName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1849)
  @PropPut
  void setSubtotalName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   */

  @DISPID(139)
  @PropGet
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(139)
  @PropPut
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DrilledDown"
   * </p>
   */

  @DISPID(1850)
  @PropGet
  boolean getDrilledDown();


  /**
   * <p>
   * Setter method for the COM property "DrilledDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1850)
  @PropPut
  void setDrilledDown(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CubeField"
   * </p>
   */

  @DISPID(1851)
  @PropGet
  com.exceljava.com4j.excel.CubeField getCubeField();


  /**
   * <p>
   * Getter method for the COM property "CurrentPageName"
   * </p>
   */

  @DISPID(1852)
  @PropGet
  java.lang.String getCurrentPageName();


  /**
   * <p>
   * Setter method for the COM property "CurrentPageName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1852)
  @PropPut
  void setCurrentPageName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "StandardFormula"
   * </p>
   */

  @DISPID(2084)
  @PropGet
  java.lang.String getStandardFormula();


  /**
   * <p>
   * Setter method for the COM property "StandardFormula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2084)
  @PropPut
  void setStandardFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "HiddenItemsList"
   * </p>
   */

  @DISPID(2139)
  @PropGet
  java.lang.Object getHiddenItemsList();


  /**
   * <p>
   * Setter method for the COM property "HiddenItemsList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2139)
  @PropPut
  void setHiddenItemsList(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DatabaseSort"
   * </p>
   */

  @DISPID(2140)
  @PropGet
  boolean getDatabaseSort();


  /**
   * <p>
   * Setter method for the COM property "DatabaseSort"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2140)
  @PropPut
  void setDatabaseSort(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IsMemberProperty"
   * </p>
   */

  @DISPID(2141)
  @PropGet
  boolean getIsMemberProperty();


  /**
   * <p>
   * Getter method for the COM property "PropertyParentField"
   * </p>
   */

  @DISPID(2142)
  @PropGet
  com.exceljava.com4j.excel.PivotField getPropertyParentField();


  /**
   * <p>
   * Getter method for the COM property "PropertyOrder"
   * </p>
   */

  @DISPID(2143)
  @PropGet
  int getPropertyOrder();


  /**
   * <p>
   * Setter method for the COM property "PropertyOrder"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2143)
  @PropPut
  void setPropertyOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableItemSelection"
   * </p>
   */

  @DISPID(2144)
  @PropGet
  boolean getEnableItemSelection();


  /**
   * <p>
   * Setter method for the COM property "EnableItemSelection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2144)
  @PropPut
  void setEnableItemSelection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CurrentPageList"
   * </p>
   */

  @DISPID(2145)
  @PropGet
  java.lang.Object getCurrentPageList();


  /**
   * <p>
   * Setter method for the COM property "CurrentPageList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2145)
  @PropPut
  void setCurrentPageList(
    java.lang.Object rhs);


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

  @DISPID(2146)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addPageItem(
    java.lang.String item);

  /**
   * @param item Mandatory java.lang.String parameter.
   * @param clearList Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2146)
  void addPageItem(
    java.lang.String item,
    @Optional java.lang.Object clearList);


  /**
   * <p>
   * Getter method for the COM property "Hidden"
   * </p>
   */

  @DISPID(268)
  @PropGet
  boolean getHidden();


  /**
   * <p>
   * Setter method for the COM property "Hidden"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(268)
  @PropPut
  void setHidden(
    boolean rhs);


  /**
   * @param field Mandatory java.lang.String parameter.
   */

  @DISPID(2580)
  void drillTo(
    java.lang.String field);


  /**
   * <p>
   * Getter method for the COM property "UseMemberPropertyAsCaption"
   * </p>
   */

  @DISPID(2581)
  @PropGet
  boolean getUseMemberPropertyAsCaption();


  /**
   * <p>
   * Setter method for the COM property "UseMemberPropertyAsCaption"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2581)
  @PropPut
  void setUseMemberPropertyAsCaption(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MemberPropertyCaption"
   * </p>
   */

  @DISPID(2582)
  @PropGet
  java.lang.String getMemberPropertyCaption();


  /**
   * <p>
   * Setter method for the COM property "MemberPropertyCaption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2582)
  @PropPut
  void setMemberPropertyCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAsTooltip"
   * </p>
   */

  @DISPID(2583)
  @PropGet
  boolean getDisplayAsTooltip();


  /**
   * <p>
   * Setter method for the COM property "DisplayAsTooltip"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2583)
  @PropPut
  void setDisplayAsTooltip(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayInReport"
   * </p>
   */

  @DISPID(2584)
  @PropGet
  boolean getDisplayInReport();


  /**
   * <p>
   * Setter method for the COM property "DisplayInReport"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2584)
  @PropPut
  void setDisplayInReport(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayAsCaption"
   * </p>
   */

  @DISPID(2585)
  @PropGet
  boolean getDisplayAsCaption();


  /**
   * <p>
   * Getter method for the COM property "LayoutCompactRow"
   * </p>
   */

  @DISPID(2586)
  @PropGet
  boolean getLayoutCompactRow();


  /**
   * <p>
   * Setter method for the COM property "LayoutCompactRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2586)
  @PropPut
  void setLayoutCompactRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IncludeNewItemsInFilter"
   * </p>
   */

  @DISPID(2587)
  @PropGet
  boolean getIncludeNewItemsInFilter();


  /**
   * <p>
   * Setter method for the COM property "IncludeNewItemsInFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2587)
  @PropPut
  void setIncludeNewItemsInFilter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleItemsList"
   * </p>
   */

  @DISPID(2588)
  @PropGet
  java.lang.Object getVisibleItemsList();


  /**
   * <p>
   * Setter method for the COM property "VisibleItemsList"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2588)
  @PropPut
  void setVisibleItemsList(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PivotFilters"
   * </p>
   */

  @DISPID(2589)
  @PropGet
  com.exceljava.com4j.excel.PivotFilters getPivotFilters();


  /**
   * <p>
   * Getter method for the COM property "AutoSortPivotLine"
   * </p>
   */

  @DISPID(2590)
  @PropGet
  com.exceljava.com4j.excel.PivotLine getAutoSortPivotLine();


  /**
   * <p>
   * Getter method for the COM property "AutoSortCustomSubtotal"
   * </p>
   */

  @DISPID(2591)
  @PropGet
  int getAutoSortCustomSubtotal();


  /**
   * <p>
   * Getter method for the COM property "ShowingInAxis"
   * </p>
   */

  @DISPID(2592)
  @PropGet
  boolean getShowingInAxis();


  /**
   * <p>
   * Getter method for the COM property "EnableMultiplePageItems"
   * </p>
   */

  @DISPID(2184)
  @PropGet
  boolean getEnableMultiplePageItems();


  /**
   * <p>
   * Setter method for the COM property "EnableMultiplePageItems"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2184)
  @PropPut
  void setEnableMultiplePageItems(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AllItemsVisible"
   * </p>
   */

  @DISPID(2593)
  @PropGet
  boolean getAllItemsVisible();


  /**
   */

  @DISPID(2594)
  void clearManualFilter();


  /**
   */

  @DISPID(2561)
  void clearAllFilters();


  /**
   */

  @DISPID(2595)
  void clearValueFilters();


  /**
   */

  @DISPID(2596)
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

  @DISPID(1514)
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

  @DISPID(1514)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void autoSort(
    int order,
    java.lang.String field,
    @Optional java.lang.Object pivotLine);

  /**
   * @param order Mandatory int parameter.
   * @param field Mandatory java.lang.String parameter.
   * @param pivotLine Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customSubtotal Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1514)
  void autoSort(
    int order,
    java.lang.String field,
    @Optional java.lang.Object pivotLine,
    @Optional java.lang.Object customSubtotal);


  /**
   * <p>
   * Getter method for the COM property "SourceCaption"
   * </p>
   */

  @DISPID(2599)
  @PropGet
  java.lang.String getSourceCaption();


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   */

  @DISPID(585)
  @PropGet
  boolean getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(585)
  @PropPut
  void setShowDetail(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RepeatLabels"
   * </p>
   */

  @DISPID(2885)
  @PropGet
  boolean getRepeatLabels();


  /**
   * <p>
   * Setter method for the COM property "RepeatLabels"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2885)
  @PropPut
  void setRepeatLabels(
    boolean rhs);


  /**
   */

  @DISPID(3165)
  void autoGroup();


  // Properties:
}
