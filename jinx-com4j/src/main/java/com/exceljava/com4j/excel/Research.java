package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Research extends Com4jObject {
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
     * <li>java.lang.Object parameter queryString is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter queryLanguage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter useSelection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter launchQuery is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * query(serviceID, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param serviceID Mandatory java.lang.String parameter.
   */

  @DISPID(2751)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object query(
    java.lang.String serviceID);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter queryLanguage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter useSelection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter launchQuery is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * query(serviceID, queryString, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param serviceID Mandatory java.lang.String parameter.
   * @param queryString Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2751)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional java.lang.Object queryString);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useSelection is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter launchQuery is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * query(serviceID, queryString, queryLanguage, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param serviceID Mandatory java.lang.String parameter.
   * @param queryString Optional parameter. Default value is com4j.Variant.getMissing()
   * @param queryLanguage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2751)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional java.lang.Object queryString,
    @Optional java.lang.Object queryLanguage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter launchQuery is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * query(serviceID, queryString, queryLanguage, useSelection, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param serviceID Mandatory java.lang.String parameter.
   * @param queryString Optional parameter. Default value is com4j.Variant.getMissing()
   * @param queryLanguage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useSelection Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2751)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional java.lang.Object queryString,
    @Optional java.lang.Object queryLanguage,
    @Optional java.lang.Object useSelection);

  /**
   * @param serviceID Mandatory java.lang.String parameter.
   * @param queryString Optional parameter. Default value is com4j.Variant.getMissing()
   * @param queryLanguage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useSelection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param launchQuery Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2751)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional java.lang.Object queryString,
    @Optional java.lang.Object queryLanguage,
    @Optional java.lang.Object useSelection,
    @Optional java.lang.Object launchQuery);


  /**
   * @param serviceID Mandatory java.lang.String parameter.
   */

  @DISPID(2757)
  boolean isResearchService(
    java.lang.String serviceID);


  /**
   * @param languageFrom Mandatory int parameter.
   * @param languageTo Mandatory int parameter.
   */

  @DISPID(2758)
  java.lang.Object setLanguagePair(
    int languageFrom,
    int languageTo);


  // Properties:
}
