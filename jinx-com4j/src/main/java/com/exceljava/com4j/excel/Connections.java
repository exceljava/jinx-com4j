package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Connections extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(2700)
  com.exceljava.com4j.excel.WorkbookConnection _AddFromFile(
    java.lang.String filename);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lCmdtype is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, description, connectionString, commandText, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection add(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.WorkbookConnection add(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText,
    @Optional java.lang.Object lCmdtype);


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel.WorkbookConnection item(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.WorkbookConnection get_Default(
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter lCmdtype is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createModelConnection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importRelationships is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(name, description, connectionString, commandText, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createModelConnection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importRelationships is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(name, description, connectionString, commandText, lCmdtype, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText,
    @Optional java.lang.Object lCmdtype);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter importRelationships is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(name, description, connectionString, commandText, lCmdtype, createModelConnection, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText,
    @Optional java.lang.Object lCmdtype,
    @Optional java.lang.Object createModelConnection);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param importRelationships Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3054)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    java.lang.Object connectionString,
    java.lang.Object commandText,
    @Optional java.lang.Object lCmdtype,
    @Optional java.lang.Object createModelConnection,
    @Optional java.lang.Object importRelationships);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createModelConnection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importRelationships is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFromFile(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(3107)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection addFromFile(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter importRelationships is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addFromFile(filename, createModelConnection, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3107)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.WorkbookConnection addFromFile(
    java.lang.String filename,
    @Optional java.lang.Object createModelConnection);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param importRelationships Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3107)
  com.exceljava.com4j.excel.WorkbookConnection addFromFile(
    java.lang.String filename,
    @Optional java.lang.Object createModelConnection,
    @Optional java.lang.Object importRelationships);


  // Properties:
}
