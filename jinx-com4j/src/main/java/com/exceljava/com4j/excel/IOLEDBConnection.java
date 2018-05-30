package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002448D-0001-0000-C000-000000000046}")
public interface IOLEDBConnection extends Com4jObject {
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
   * Getter method for the COM property "ADOConnection"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getADOConnection();


  /**
   * <p>
   * Getter method for the COM property "BackgroundQuery"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(11)
  boolean getBackgroundQuery();


  /**
   * <p>
   * Setter method for the COM property "BackgroundQuery"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(12)
  void setBackgroundQuery(
    boolean rhs);


  /**
   */

  @VTID(13)
  void cancelRefresh();


  /**
   * <p>
   * Getter method for the COM property "CommandText"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCommandText();


  /**
   * <p>
   * Setter method for the COM property "CommandText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setCommandText(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCmdType
   */

  @VTID(16)
  com.exceljava.com4j.excel.XlCmdType getCommandType();


  /**
   * <p>
   * Setter method for the COM property "CommandType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCmdType parameter.
   */

  @VTID(17)
  void setCommandType(
    com.exceljava.com4j.excel.XlCmdType rhs);


  /**
   * <p>
   * Getter method for the COM property "Connection"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getConnection();


  /**
   * <p>
   * Setter method for the COM property "Connection"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(19)
  void setConnection(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableRefresh"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getEnableRefresh();


  /**
   * <p>
   * Setter method for the COM property "EnableRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setEnableRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LocalConnection"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLocalConnection();


  /**
   * <p>
   * Setter method for the COM property "LocalConnection"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(23)
  void setLocalConnection(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "MaintainConnection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getMaintainConnection();


  /**
   * <p>
   * Setter method for the COM property "MaintainConnection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setMaintainConnection(
    boolean rhs);


  /**
   */

  @VTID(26)
  void makeConnection();


  /**
   */

  @VTID(27)
  void refresh();


  /**
   * <p>
   * Getter method for the COM property "RefreshDate"
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @VTID(28)
  java.util.Date getRefreshDate();


  /**
   * <p>
   * Getter method for the COM property "Refreshing"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getRefreshing();


  /**
   * <p>
   * Getter method for the COM property "RefreshOnFileOpen"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getRefreshOnFileOpen();


  /**
   * <p>
   * Setter method for the COM property "RefreshOnFileOpen"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setRefreshOnFileOpen(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RefreshPeriod"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(32)
  int getRefreshPeriod();


  /**
   * <p>
   * Setter method for the COM property "RefreshPeriod"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(33)
  void setRefreshPeriod(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "RobustConnect"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlRobustConnect
   */

  @VTID(34)
  com.exceljava.com4j.excel.XlRobustConnect getRobustConnect();


  /**
   * <p>
   * Setter method for the COM property "RobustConnect"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRobustConnect parameter.
   */

  @VTID(35)
  void setRobustConnect(
    com.exceljava.com4j.excel.XlRobustConnect rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter keywords is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAsODC(odcFileName, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param odcFileName Mandatory java.lang.String parameter.
   */

  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void saveAsODC(
    java.lang.String odcFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter keywords is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAsODC(odcFileName, description, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param odcFileName Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);

  /**
   * @param odcFileName Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param keywords Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(36)
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object keywords);


  /**
   * <p>
   * Getter method for the COM property "SavePassword"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(37)
  boolean getSavePassword();


  /**
   * <p>
   * Setter method for the COM property "SavePassword"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(38)
  void setSavePassword(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceConnectionFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(39)
  java.lang.String getSourceConnectionFile();


  /**
   * <p>
   * Setter method for the COM property "SourceConnectionFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(40)
  void setSourceConnectionFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceDataFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(41)
  java.lang.String getSourceDataFile();


  /**
   * <p>
   * Setter method for the COM property "SourceDataFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(42)
  void setSourceDataFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OLAP"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(43)
  boolean getOLAP();


  /**
   * <p>
   * Getter method for the COM property "UseLocalConnection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(44)
  boolean getUseLocalConnection();


  /**
   * <p>
   * Setter method for the COM property "UseLocalConnection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(45)
  void setUseLocalConnection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MaxDrillthroughRecords"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(46)
  int getMaxDrillthroughRecords();


  /**
   * <p>
   * Setter method for the COM property "MaxDrillthroughRecords"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(47)
  void setMaxDrillthroughRecords(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "IsConnected"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(48)
  boolean getIsConnected();


  /**
   * <p>
   * Getter method for the COM property "ServerCredentialsMethod"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCredentialsMethod
   */

  @VTID(49)
  com.exceljava.com4j.excel.XlCredentialsMethod getServerCredentialsMethod();


  /**
   * <p>
   * Setter method for the COM property "ServerCredentialsMethod"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCredentialsMethod parameter.
   */

  @VTID(50)
  void setServerCredentialsMethod(
    com.exceljava.com4j.excel.XlCredentialsMethod rhs);


  /**
   * <p>
   * Getter method for the COM property "ServerSSOApplicationID"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(51)
  java.lang.String getServerSSOApplicationID();


  /**
   * <p>
   * Setter method for the COM property "ServerSSOApplicationID"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(52)
  void setServerSSOApplicationID(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "AlwaysUseConnectionFile"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(53)
  boolean getAlwaysUseConnectionFile();


  /**
   * <p>
   * Setter method for the COM property "AlwaysUseConnectionFile"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(54)
  void setAlwaysUseConnectionFile(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ServerFillColor"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(55)
  boolean getServerFillColor();


  /**
   * <p>
   * Setter method for the COM property "ServerFillColor"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(56)
  void setServerFillColor(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ServerFontStyle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(57)
  boolean getServerFontStyle();


  /**
   * <p>
   * Setter method for the COM property "ServerFontStyle"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(58)
  void setServerFontStyle(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ServerNumberFormat"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(59)
  boolean getServerNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "ServerNumberFormat"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(60)
  void setServerNumberFormat(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ServerTextColor"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(61)
  boolean getServerTextColor();


  /**
   * <p>
   * Setter method for the COM property "ServerTextColor"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(62)
  void setServerTextColor(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RetrieveInOfficeUILang"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(63)
  boolean getRetrieveInOfficeUILang();


  /**
   * <p>
   * Setter method for the COM property "RetrieveInOfficeUILang"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(64)
  void setRetrieveInOfficeUILang(
    boolean rhs);


  /**
   */

  @VTID(65)
  void reconnect();


  /**
   * <p>
   * Getter method for the COM property "CalculatedMembers"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMembers
   */

  @VTID(66)
  com.exceljava.com4j.excel.CalculatedMembers getCalculatedMembers();


  /**
   * <p>
   * Getter method for the COM property "LocaleID"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(67)
  int getLocaleID();


  /**
   * <p>
   * Setter method for the COM property "LocaleID"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(68)
  void setLocaleID(
    int rhs);


  // Properties:
}
