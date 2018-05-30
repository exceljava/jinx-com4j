package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024484-0001-0000-C000-000000000046}")
public interface IPivotFilters extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.excel.PivotFilter get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(12)
  com.exceljava.com4j.excel.PivotFilter getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlPivotFilterType parameter.
   * @param dataField Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param memberPropertyField Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(14)
  com.exceljava.com4j.excel.PivotFilter add(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object memberPropertyField);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 9}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 9}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 9}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 9}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 9}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 9}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 9}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 9}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object memberPropertyField);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFilter
   */

  @VTID(15)
  com.exceljava.com4j.excel.PivotFilter add2(
    com.exceljava.com4j.excel.XlPivotFilterType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object memberPropertyField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object wholeDayFilter);


  // Properties:
}
