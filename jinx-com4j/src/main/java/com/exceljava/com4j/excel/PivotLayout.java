package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotLayout extends Com4jObject {
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
   * Getter method for the COM property "ColumnFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getColumnFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(713)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getColumnFields();

  /**
   * <p>
   * Getter method for the COM property "ColumnFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(713)
  @PropGet
  com4j.Com4jObject getColumnFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "DataFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDataFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(715)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getDataFields();

  /**
   * <p>
   * Getter method for the COM property "DataFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(715)
  @PropGet
  com4j.Com4jObject getDataFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "PageFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPageFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(714)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getPageFields();

  /**
   * <p>
   * Getter method for the COM property "PageFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(714)
  @PropGet
  com4j.Com4jObject getPageFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "RowFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRowFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(712)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getRowFields();

  /**
   * <p>
   * Getter method for the COM property "RowFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(712)
  @PropGet
  com4j.Com4jObject getRowFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "HiddenFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHiddenFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(711)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getHiddenFields();

  /**
   * <p>
   * Getter method for the COM property "HiddenFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(711)
  @PropGet
  com4j.Com4jObject getHiddenFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "VisibleFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getVisibleFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(710)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getVisibleFields();

  /**
   * <p>
   * Getter method for the COM property "VisibleFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(710)
  @PropGet
  com4j.Com4jObject getVisibleFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "PivotFields"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPivotFields(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(718)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject getPivotFields();

  /**
   * <p>
   * Getter method for the COM property "PivotFields"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(718)
  @PropGet
  com4j.Com4jObject getPivotFields(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "CubeFields"
   * </p>
   */

  @DISPID(1839)
  @PropGet
  com.exceljava.com4j.excel.CubeFields getCubeFields();


  /**
   * <p>
   * Getter method for the COM property "PivotCache"
   * </p>
   */

  @DISPID(1496)
  @PropGet
  com.exceljava.com4j.excel.PivotCache getPivotCache();


  /**
   * <p>
   * Getter method for the COM property "PivotTable"
   * </p>
   */

  @DISPID(716)
  @PropGet
  com.exceljava.com4j.excel.PivotTable getPivotTable();


  /**
   * <p>
   * Getter method for the COM property "InnerDetail"
   * </p>
   */

  @DISPID(698)
  @PropGet
  java.lang.String getInnerDetail();


  /**
   * <p>
   * Setter method for the COM property "InnerDetail"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(698)
  @PropPut
  void setInnerDetail(
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter appendField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFields(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(708)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void addFields();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter appendField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFields(rowFields, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowFields Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(708)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void addFields(
    @Optional java.lang.Object rowFields);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageFields is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter appendField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFields(rowFields, columnFields, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnFields Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(708)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void addFields(
    @Optional java.lang.Object rowFields,
    @Optional java.lang.Object columnFields);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter appendField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFields(rowFields, columnFields, pageFields, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFields Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(708)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addFields(
    @Optional java.lang.Object rowFields,
    @Optional java.lang.Object columnFields,
    @Optional java.lang.Object pageFields);

  /**
   * @param rowFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFields Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appendField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(708)
  void addFields(
    @Optional java.lang.Object rowFields,
    @Optional java.lang.Object columnFields,
    @Optional java.lang.Object pageFields,
    @Optional java.lang.Object appendField);


  // Properties:
}
