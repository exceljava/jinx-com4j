package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotFilters extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.PivotFilter get_Default(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.PivotFilter getItem(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, value1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, value1, value2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, value1, value2, order, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, value1, value2, order, name, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, dataField, value1, value2, order, name, description, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name,
    @Optional java.lang.Object description);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param memberPropertyField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name,
    @Optional java.lang.Object description,
    @Optional java.lang.Object memberPropertyField);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, value2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, value2, order, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, value2, order, name, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter memberPropertyField is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, value2, order, name, description, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name,
    @Optional java.lang.Object description);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter wholeDayFilter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(type, dataField, value1, value2, order, name, description, memberPropertyField, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param memberPropertyField Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name,
    @Optional java.lang.Object description,
    @Optional java.lang.Object memberPropertyField);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param memberPropertyField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param wholeDayFilter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional java.lang.Object dataField,
    @Optional java.lang.Object value1,
    @Optional java.lang.Object value2,
    @Optional java.lang.Object order,
    @Optional java.lang.Object name,
    @Optional java.lang.Object description,
    @Optional java.lang.Object memberPropertyField,
    @Optional java.lang.Object wholeDayFilter);


  // Properties:
}
