package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface XmlMap extends Com4jObject {
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
   * Getter method for the COM property "IsExportable"
   * </p>
   */

  @DISPID(2334)
  @PropGet
  boolean getIsExportable();


  /**
   * <p>
   * Getter method for the COM property "ShowImportExportValidationErrors"
   * </p>
   */

  @DISPID(2335)
  @PropGet
  boolean getShowImportExportValidationErrors();


  /**
   * <p>
   * Setter method for the COM property "ShowImportExportValidationErrors"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2335)
  @PropPut
  void setShowImportExportValidationErrors(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveDataSourceDefinition"
   * </p>
   */

  @DISPID(2336)
  @PropGet
  boolean getSaveDataSourceDefinition();


  /**
   * <p>
   * Setter method for the COM property "SaveDataSourceDefinition"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2336)
  @PropPut
  void setSaveDataSourceDefinition(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AdjustColumnWidth"
   * </p>
   */

  @DISPID(1868)
  @PropGet
  boolean getAdjustColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "AdjustColumnWidth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1868)
  @PropPut
  void setAdjustColumnWidth(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveColumnFilter"
   * </p>
   */

  @DISPID(2337)
  @PropGet
  boolean getPreserveColumnFilter();


  /**
   * <p>
   * Setter method for the COM property "PreserveColumnFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2337)
  @PropPut
  void setPreserveColumnFilter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveNumberFormatting"
   * </p>
   */

  @DISPID(2338)
  @PropGet
  boolean getPreserveNumberFormatting();


  /**
   * <p>
   * Setter method for the COM property "PreserveNumberFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2338)
  @PropPut
  void setPreserveNumberFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AppendOnImport"
   * </p>
   */

  @DISPID(2339)
  @PropGet
  boolean getAppendOnImport();


  /**
   * <p>
   * Setter method for the COM property "AppendOnImport"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2339)
  @PropPut
  void setAppendOnImport(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RootElementName"
   * </p>
   */

  @DISPID(2340)
  @PropGet
  java.lang.String getRootElementName();


  /**
   * <p>
   * Getter method for the COM property "RootElementNamespace"
   * </p>
   */

  @DISPID(2341)
  @PropGet
  com.exceljava.com4j.excel.XmlNamespace getRootElementNamespace();


  /**
   * <p>
   * Getter method for the COM property "Schemas"
   * </p>
   */

  @DISPID(2342)
  @PropGet
  com.exceljava.com4j.excel.XmlSchemas getSchemas();


  /**
   * <p>
   * Getter method for the COM property "DataBinding"
   * </p>
   */

  @DISPID(2343)
  @PropGet
  com.exceljava.com4j.excel.XmlDataBinding getDataBinding();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _import(url, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param url Mandatory java.lang.String parameter.
   */

  @DISPID(917)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.XlXmlImportResult _import(
    java.lang.String url);

  /**
   * @param url Mandatory java.lang.String parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(917)
  com.exceljava.com4j.excel.XlXmlImportResult _import(
    java.lang.String url,
    @Optional java.lang.Object overwrite);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * importXml(xmlData, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xmlData Mandatory java.lang.String parameter.
   */

  @DISPID(2344)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.XlXmlImportResult importXml(
    java.lang.String xmlData);

  /**
   * @param xmlData Mandatory java.lang.String parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2344)
  com.exceljava.com4j.excel.XlXmlImportResult importXml(
    java.lang.String xmlData,
    @Optional java.lang.Object overwrite);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * export(url, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param url Mandatory java.lang.String parameter.
   */

  @DISPID(1414)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.XlXmlExportResult export(
    java.lang.String url);

  /**
   * @param url Mandatory java.lang.String parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1414)
  com.exceljava.com4j.excel.XlXmlExportResult export(
    java.lang.String url,
    @Optional java.lang.Object overwrite);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlExportResult
   */

  @DISPID(2346)
  com.exceljava.com4j.excel.XlXmlExportResult exportXml();


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   */

  @DISPID(2544)
  @PropGet
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  // Properties:
}
