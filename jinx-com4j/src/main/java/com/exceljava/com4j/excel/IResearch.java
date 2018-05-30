package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244AC-0001-0000-C000-000000000046}")
public interface IResearch extends Com4jObject {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryString);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryString,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryLanguage);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryString,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryLanguage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useSelection);

  /**
   * @param serviceID Mandatory java.lang.String parameter.
   * @param queryString Optional parameter. Default value is com4j.Variant.getMissing()
   * @param queryLanguage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param useSelection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param launchQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object query(
    java.lang.String serviceID,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryString,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object queryLanguage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useSelection,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object launchQuery);


  /**
   * @param serviceID Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean isResearchService(
    java.lang.String serviceID);


  /**
   * @param languageFrom Mandatory int parameter.
   * @param languageTo Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object setLanguagePair(
    int languageFrom,
    int languageTo);


  // Properties:
}
