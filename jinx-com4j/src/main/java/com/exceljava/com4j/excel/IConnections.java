package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024486-0001-0000-C000-000000000046}")
public interface IConnections extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(11)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.WorkbookConnection add(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(12)
  com.exceljava.com4j.excel.WorkbookConnection add(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lCmdtype);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(13)
  com.exceljava.com4j.excel.WorkbookConnection item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(14)
  @DefaultMethod
  com.exceljava.com4j.excel.WorkbookConnection get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(15)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lCmdtype);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lCmdtype,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createModelConnection);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param connectionString Mandatory java.lang.Object parameter.
   * @param commandText Mandatory java.lang.Object parameter.
   * @param lCmdtype Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param importRelationships Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(16)
  com.exceljava.com4j.excel.WorkbookConnection add2(
    java.lang.String name,
    java.lang.String description,
    @MarshalAs(NativeType.VARIANT) java.lang.Object connectionString,
    @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object lCmdtype,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createModelConnection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object importRelationships);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.WorkbookConnection addFromFile(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createModelConnection);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param createModelConnection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param importRelationships Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(17)
  com.exceljava.com4j.excel.WorkbookConnection addFromFile(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createModelConnection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object importRelationships);


  // Properties:
}
