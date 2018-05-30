package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208DA-0000-0000-C000-000000000046}")
public interface _Workbook extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AcceptLabelsInFormulas"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1441) //= 0x5a1. The runtime will prefer the VTID if present
  @VTID(10)
  boolean getAcceptLabelsInFormulas();


  /**
   * <p>
   * Setter method for the COM property "AcceptLabelsInFormulas"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1441) //= 0x5a1. The runtime will prefer the VTID if present
  @VTID(11)
  void setAcceptLabelsInFormulas(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * activate(1033);
   * </code>
   * </p>
   */

  @DISPID(304) //= 0x130. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void activate();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(304) //= 0x130. The runtime will prefer the VTID if present
  @VTID(12)
  void activate(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ActiveChart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @DISPID(183) //= 0xb7. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.excel._Chart getActiveChart();


  /**
   * <p>
   * Getter method for the COM property "ActiveSheet"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(307) //= 0x133. The runtime will prefer the VTID if present
  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getActiveSheet();


  /**
   * <p>
   * Getter method for the COM property "Author"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAuthor(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(574) //= 0x23e. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getAuthor();

  /**
   * <p>
   * Getter method for the COM property "Author"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(574) //= 0x23e. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getAuthor(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Author"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setAuthor(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(574) //= 0x23e. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setAuthor(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "Author"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(574) //= 0x23e. The runtime will prefer the VTID if present
  @VTID(16)
  void setAuthor(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoUpdateFrequency"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1442) //= 0x5a2. The runtime will prefer the VTID if present
  @VTID(17)
  int getAutoUpdateFrequency();


  /**
   * <p>
   * Setter method for the COM property "AutoUpdateFrequency"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1442) //= 0x5a2. The runtime will prefer the VTID if present
  @VTID(18)
  void setAutoUpdateFrequency(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoUpdateSaveChanges"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1443) //= 0x5a3. The runtime will prefer the VTID if present
  @VTID(19)
  boolean getAutoUpdateSaveChanges();


  /**
   * <p>
   * Setter method for the COM property "AutoUpdateSaveChanges"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1443) //= 0x5a3. The runtime will prefer the VTID if present
  @VTID(20)
  void setAutoUpdateSaveChanges(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ChangeHistoryDuration"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1444) //= 0x5a4. The runtime will prefer the VTID if present
  @VTID(21)
  int getChangeHistoryDuration();


  /**
   * <p>
   * Setter method for the COM property "ChangeHistoryDuration"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1444) //= 0x5a4. The runtime will prefer the VTID if present
  @VTID(22)
  void setChangeHistoryDuration(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "BuiltinDocumentProperties"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1176) //= 0x498. The runtime will prefer the VTID if present
  @VTID(23)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getBuiltinDocumentProperties();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writePassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * changeFileAccess(mode, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param mode Mandatory com.exceljava.com4j.excel.XlFileAccess parameter.
   */

  @DISPID(989) //= 0x3dd. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void changeFileAccess(
    com.exceljava.com4j.excel.XlFileAccess mode);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * changeFileAccess(mode, writePassword, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param mode Mandatory com.exceljava.com4j.excel.XlFileAccess parameter.
   * @param writePassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(989) //= 0x3dd. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void changeFileAccess(
    com.exceljava.com4j.excel.XlFileAccess mode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writePassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * changeFileAccess(mode, writePassword, notify, 1033);
   * </code>
   * </p>
   * @param mode Mandatory com.exceljava.com4j.excel.XlFileAccess parameter.
   * @param writePassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(989) //= 0x3dd. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void changeFileAccess(
    com.exceljava.com4j.excel.XlFileAccess mode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writePassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify);

  /**
   * @param mode Mandatory com.exceljava.com4j.excel.XlFileAccess parameter.
   * @param writePassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(989) //= 0x3dd. The runtime will prefer the VTID if present
  @VTID(24)
  void changeFileAccess(
    com.exceljava.com4j.excel.XlFileAccess mode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writePassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlLinkType parameter type is set to 1</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * changeLink(name, newName, 1, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param newName Mandatory java.lang.String parameter.
   */

  @DISPID(802) //= 0x322. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.excel.XlLinkType.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "1033"})
  void changeLink(
    java.lang.String name,
    java.lang.String newName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * changeLink(name, newName, type, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param newName Mandatory java.lang.String parameter.
   * @param type Optional parameter. Default value is 1
   */

  @DISPID(802) //= 0x322. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void changeLink(
    java.lang.String name,
    java.lang.String newName,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlLinkType type);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param newName Mandatory java.lang.String parameter.
   * @param type Optional parameter. Default value is 1
   * @param lcid Mandatory int parameter.
   */

  @DISPID(802) //= 0x322. The runtime will prefer the VTID if present
  @VTID(25)
  void changeLink(
    java.lang.String name,
    java.lang.String newName,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlLinkType type,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Charts"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.excel.Sheets getCharts();


  @VTID(26)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getCharts(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveChanges is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void close();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(saveChanges, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(saveChanges, filename, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(saveChanges, filename, routeWorkbook, 1033);
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param routeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object routeWorkbook);

  /**
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param routeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(27)
  void close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object routeWorkbook,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CodeName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1373) //= 0x55d. The runtime will prefer the VTID if present
  @VTID(28)
  java.lang.String getCodeName();


  /**
   * <p>
   * Getter method for the COM property "_CodeName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-2147418112) //= 0x80010000. The runtime will prefer the VTID if present
  @VTID(29)
  java.lang.String get_CodeName();


  /**
   * <p>
   * Setter method for the COM property "_CodeName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(-2147418112) //= 0x80010000. The runtime will prefer the VTID if present
  @VTID(30)
  void set_CodeName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Colors"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getColors(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getColors();

  /**
   * <p>
   * Getter method for the COM property "Colors"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getColors(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getColors(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Colors"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(31)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getColors(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Colors"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setColors(com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void setColors(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Colors"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setColors(index, 1033, rhs);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setColors(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Colors"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(286) //= 0x11e. The runtime will prefer the VTID if present
  @VTID(32)
  void setColors(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CommandBars
   */

  @DISPID(1439) //= 0x59f. The runtime will prefer the VTID if present
  @VTID(33)
  com.exceljava.com4j.office._CommandBars getCommandBars();


  /**
   * <p>
   * Getter method for the COM property "Comments"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getComments(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(575) //= 0x23f. The runtime will prefer the VTID if present
  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getComments();

  /**
   * <p>
   * Getter method for the COM property "Comments"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(575) //= 0x23f. The runtime will prefer the VTID if present
  @VTID(34)
  java.lang.String getComments(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Comments"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setComments(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(575) //= 0x23f. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setComments(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "Comments"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(575) //= 0x23f. The runtime will prefer the VTID if present
  @VTID(35)
  void setComments(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ConflictResolution"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSaveConflictResolution
   */

  @DISPID(1175) //= 0x497. The runtime will prefer the VTID if present
  @VTID(36)
  com.exceljava.com4j.excel.XlSaveConflictResolution getConflictResolution();


  /**
   * <p>
   * Setter method for the COM property "ConflictResolution"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSaveConflictResolution parameter.
   */

  @DISPID(1175) //= 0x497. The runtime will prefer the VTID if present
  @VTID(37)
  void setConflictResolution(
    com.exceljava.com4j.excel.XlSaveConflictResolution rhs);


  /**
   * <p>
   * Getter method for the COM property "Container"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1190) //= 0x4a6. The runtime will prefer the VTID if present
  @VTID(38)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getContainer();


  /**
   * <p>
   * Getter method for the COM property "CreateBackup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCreateBackup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(287) //= 0x11f. The runtime will prefer the VTID if present
  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getCreateBackup();

  /**
   * <p>
   * Getter method for the COM property "CreateBackup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(287) //= 0x11f. The runtime will prefer the VTID if present
  @VTID(39)
  boolean getCreateBackup(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CustomDocumentProperties"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1177) //= 0x499. The runtime will prefer the VTID if present
  @VTID(40)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getCustomDocumentProperties();


  /**
   * <p>
   * Getter method for the COM property "Date1904"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDate1904(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(403) //= 0x193. The runtime will prefer the VTID if present
  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getDate1904();

  /**
   * <p>
   * Getter method for the COM property "Date1904"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(403) //= 0x193. The runtime will prefer the VTID if present
  @VTID(41)
  boolean getDate1904(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Date1904"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDate1904(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(403) //= 0x193. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDate1904(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "Date1904"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(403) //= 0x193. The runtime will prefer the VTID if present
  @VTID(42)
  void setDate1904(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * deleteNumberFormat(numberFormat, 1033);
   * </code>
   * </p>
   * @param numberFormat Mandatory java.lang.String parameter.
   */

  @DISPID(397) //= 0x18d. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void deleteNumberFormat(
    java.lang.String numberFormat);

  /**
   * @param numberFormat Mandatory java.lang.String parameter.
   * @param lcid Mandatory int parameter.
   */

  @DISPID(397) //= 0x18d. The runtime will prefer the VTID if present
  @VTID(43)
  void deleteNumberFormat(
    java.lang.String numberFormat,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "DialogSheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(764) //= 0x2fc. The runtime will prefer the VTID if present
  @VTID(44)
  com.exceljava.com4j.excel.Sheets getDialogSheets();


  @VTID(44)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getDialogSheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "DisplayDrawingObjects"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDisplayDrawingObjects(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDisplayDrawingObjects
   */

  @DISPID(404) //= 0x194. The runtime will prefer the VTID if present
  @VTID(45)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlDisplayDrawingObjects getDisplayDrawingObjects();

  /**
   * <p>
   * Getter method for the COM property "DisplayDrawingObjects"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDisplayDrawingObjects
   */

  @DISPID(404) //= 0x194. The runtime will prefer the VTID if present
  @VTID(45)
  com.exceljava.com4j.excel.XlDisplayDrawingObjects getDisplayDrawingObjects(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "DisplayDrawingObjects"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDisplayDrawingObjects(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDisplayDrawingObjects parameter.
   */

  @DISPID(404) //= 0x194. The runtime will prefer the VTID if present
  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDisplayDrawingObjects(
    com.exceljava.com4j.excel.XlDisplayDrawingObjects rhs);

  /**
   * <p>
   * Setter method for the COM property "DisplayDrawingObjects"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDisplayDrawingObjects parameter.
   */

  @DISPID(404) //= 0x194. The runtime will prefer the VTID if present
  @VTID(46)
  void setDisplayDrawingObjects(
    @LCID int lcid,
    com.exceljava.com4j.excel.XlDisplayDrawingObjects rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exclusiveAccess(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1168) //= 0x490. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean exclusiveAccess();

  /**
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1168) //= 0x490. The runtime will prefer the VTID if present
  @VTID(47)
  boolean exclusiveAccess(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "FileFormat"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFileFormat(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlFileFormat
   */

  @DISPID(288) //= 0x120. The runtime will prefer the VTID if present
  @VTID(48)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlFileFormat getFileFormat();

  /**
   * <p>
   * Getter method for the COM property "FileFormat"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlFileFormat
   */

  @DISPID(288) //= 0x120. The runtime will prefer the VTID if present
  @VTID(48)
  com.exceljava.com4j.excel.XlFileFormat getFileFormat(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * forwardMailer(1033);
   * </code>
   * </p>
   */

  @DISPID(973) //= 0x3cd. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void forwardMailer();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(973) //= 0x3cd. The runtime will prefer the VTID if present
  @VTID(49)
  void forwardMailer(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "FullName"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFullName(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(289) //= 0x121. The runtime will prefer the VTID if present
  @VTID(50)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getFullName();

  /**
   * <p>
   * Getter method for the COM property "FullName"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(289) //= 0x121. The runtime will prefer the VTID if present
  @VTID(50)
  java.lang.String getFullName(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "HasMailer"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasMailer(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(976) //= 0x3d0. The runtime will prefer the VTID if present
  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getHasMailer();

  /**
   * <p>
   * Getter method for the COM property "HasMailer"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(976) //= 0x3d0. The runtime will prefer the VTID if present
  @VTID(51)
  boolean getHasMailer(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HasMailer"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasMailer(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(976) //= 0x3d0. The runtime will prefer the VTID if present
  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHasMailer(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "HasMailer"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(976) //= 0x3d0. The runtime will prefer the VTID if present
  @VTID(52)
  void setHasMailer(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasPassword"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasPassword(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(290) //= 0x122. The runtime will prefer the VTID if present
  @VTID(53)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getHasPassword();

  /**
   * <p>
   * Getter method for the COM property "HasPassword"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(290) //= 0x122. The runtime will prefer the VTID if present
  @VTID(53)
  boolean getHasPassword(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "HasRoutingSlip"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasRoutingSlip(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(950) //= 0x3b6. The runtime will prefer the VTID if present
  @VTID(54)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getHasRoutingSlip();

  /**
   * <p>
   * Getter method for the COM property "HasRoutingSlip"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(950) //= 0x3b6. The runtime will prefer the VTID if present
  @VTID(54)
  boolean getHasRoutingSlip(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HasRoutingSlip"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasRoutingSlip(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(950) //= 0x3b6. The runtime will prefer the VTID if present
  @VTID(55)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHasRoutingSlip(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "HasRoutingSlip"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(950) //= 0x3b6. The runtime will prefer the VTID if present
  @VTID(55)
  void setHasRoutingSlip(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IsAddin"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1445) //= 0x5a5. The runtime will prefer the VTID if present
  @VTID(56)
  boolean getIsAddin();


  /**
   * <p>
   * Setter method for the COM property "IsAddin"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1445) //= 0x5a5. The runtime will prefer the VTID if present
  @VTID(57)
  void setIsAddin(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Keywords"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getKeywords(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(577) //= 0x241. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getKeywords();

  /**
   * <p>
   * Getter method for the COM property "Keywords"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(577) //= 0x241. The runtime will prefer the VTID if present
  @VTID(58)
  java.lang.String getKeywords(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Keywords"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setKeywords(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(577) //= 0x241. The runtime will prefer the VTID if present
  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setKeywords(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "Keywords"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(577) //= 0x241. The runtime will prefer the VTID if present
  @VTID(59)
  void setKeywords(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editionRef is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * linkInfo(name, linkInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkInfo Mandatory com.exceljava.com4j.excel.XlLinkInfo parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(807) //= 0x327. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object linkInfo(
    java.lang.String name,
    com.exceljava.com4j.excel.XlLinkInfo linkInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter editionRef is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * linkInfo(name, linkInfo, type, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkInfo Mandatory com.exceljava.com4j.excel.XlLinkInfo parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(807) //= 0x327. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object linkInfo(
    java.lang.String name,
    com.exceljava.com4j.excel.XlLinkInfo linkInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * linkInfo(name, linkInfo, type, editionRef, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkInfo Mandatory com.exceljava.com4j.excel.XlLinkInfo parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editionRef Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(807) //= 0x327. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object linkInfo(
    java.lang.String name,
    com.exceljava.com4j.excel.XlLinkInfo linkInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editionRef);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param linkInfo Mandatory com.exceljava.com4j.excel.XlLinkInfo parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editionRef Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(807) //= 0x327. The runtime will prefer the VTID if present
  @VTID(60)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object linkInfo(
    java.lang.String name,
    com.exceljava.com4j.excel.XlLinkInfo linkInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editionRef,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * linkSources(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(808) //= 0x328. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object linkSources();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * linkSources(type, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(808) //= 0x328. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object linkSources(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(808) //= 0x328. The runtime will prefer the VTID if present
  @VTID(61)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object linkSources(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Mailer"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Mailer
   */

  @DISPID(979) //= 0x3d3. The runtime will prefer the VTID if present
  @VTID(62)
  com.exceljava.com4j.excel.Mailer getMailer();


  /**
   * @param filename Mandatory java.lang.Object parameter.
   */

  @DISPID(1446) //= 0x5a6. The runtime will prefer the VTID if present
  @VTID(63)
  void mergeWorkbook(
    @MarshalAs(NativeType.VARIANT) java.lang.Object filename);


  /**
   * <p>
   * Getter method for the COM property "Modules"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(582) //= 0x246. The runtime will prefer the VTID if present
  @VTID(64)
  com.exceljava.com4j.excel.Sheets getModules();


  @VTID(64)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getModules(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "MultiUserEditing"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getMultiUserEditing(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1169) //= 0x491. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getMultiUserEditing();

  /**
   * <p>
   * Getter method for the COM property "MultiUserEditing"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1169) //= 0x491. The runtime will prefer the VTID if present
  @VTID(65)
  boolean getMultiUserEditing(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(66)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "Names"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Names
   */

  @DISPID(442) //= 0x1ba. The runtime will prefer the VTID if present
  @VTID(67)
  com.exceljava.com4j.excel.Names getNames();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * newWindow(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Window
   */

  @DISPID(280) //= 0x118. The runtime will prefer the VTID if present
  @VTID(68)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Window newWindow();

  /**
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Window
   */

  @DISPID(280) //= 0x118. The runtime will prefer the VTID if present
  @VTID(68)
  com.exceljava.com4j.excel.Window newWindow(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "OnSave"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnSave(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1178) //= 0x49a. The runtime will prefer the VTID if present
  @VTID(69)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnSave();

  /**
   * <p>
   * Getter method for the COM property "OnSave"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1178) //= 0x49a. The runtime will prefer the VTID if present
  @VTID(69)
  java.lang.String getOnSave(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnSave"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnSave(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1178) //= 0x49a. The runtime will prefer the VTID if present
  @VTID(70)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnSave(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnSave"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1178) //= 0x49a. The runtime will prefer the VTID if present
  @VTID(70)
  void setOnSave(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnSheetActivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnSheetActivate(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(71)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnSheetActivate();

  /**
   * <p>
   * Getter method for the COM property "OnSheetActivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(71)
  java.lang.String getOnSheetActivate(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnSheetActivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnSheetActivate(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(72)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnSheetActivate(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnSheetActivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(72)
  void setOnSheetActivate(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnSheetDeactivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnSheetDeactivate(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnSheetDeactivate();

  /**
   * <p>
   * Getter method for the COM property "OnSheetDeactivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(73)
  java.lang.String getOnSheetDeactivate(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnSheetDeactivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnSheetDeactivate(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnSheetDeactivate(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnSheetDeactivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(74)
  void setOnSheetDeactivate(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openLinks(name, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   */

  @DISPID(803) //= 0x323. The runtime will prefer the VTID if present
  @VTID(75)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void openLinks(
    java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openLinks(name, readOnly, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(803) //= 0x323. The runtime will prefer the VTID if present
  @VTID(75)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void openLinks(
    java.lang.String name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openLinks(name, readOnly, type, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(803) //= 0x323. The runtime will prefer the VTID if present
  @VTID(75)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void openLinks(
    java.lang.String name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(803) //= 0x323. The runtime will prefer the VTID if present
  @VTID(75)
  void openLinks(
    java.lang.String name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Path"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPath(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(291) //= 0x123. The runtime will prefer the VTID if present
  @VTID(76)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getPath();

  /**
   * <p>
   * Getter method for the COM property "Path"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(291) //= 0x123. The runtime will prefer the VTID if present
  @VTID(76)
  java.lang.String getPath(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "PersonalViewListSettings"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1447) //= 0x5a7. The runtime will prefer the VTID if present
  @VTID(77)
  boolean getPersonalViewListSettings();


  /**
   * <p>
   * Setter method for the COM property "PersonalViewListSettings"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1447) //= 0x5a7. The runtime will prefer the VTID if present
  @VTID(78)
  void setPersonalViewListSettings(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PersonalViewPrintSettings"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1448) //= 0x5a8. The runtime will prefer the VTID if present
  @VTID(79)
  boolean getPersonalViewPrintSettings();


  /**
   * <p>
   * Setter method for the COM property "PersonalViewPrintSettings"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1448) //= 0x5a8. The runtime will prefer the VTID if present
  @VTID(80)
  void setPersonalViewPrintSettings(
    boolean rhs);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCaches
   */

  @DISPID(1449) //= 0x5a9. The runtime will prefer the VTID if present
  @VTID(81)
  com.exceljava.com4j.excel.PivotCaches pivotCaches();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * post(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1166) //= 0x48e. The runtime will prefer the VTID if present
  @VTID(82)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void post();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * post(destName, 1033);
   * </code>
   * </p>
   * @param destName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1166) //= 0x48e. The runtime will prefer the VTID if present
  @VTID(82)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void post(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destName);

  /**
   * @param destName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1166) //= 0x48e. The runtime will prefer the VTID if present
  @VTID(82)
  void post(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destName,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "PrecisionAsDisplayed"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPrecisionAsDisplayed(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(405) //= 0x195. The runtime will prefer the VTID if present
  @VTID(83)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getPrecisionAsDisplayed();

  /**
   * <p>
   * Getter method for the COM property "PrecisionAsDisplayed"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(405) //= 0x195. The runtime will prefer the VTID if present
  @VTID(83)
  boolean getPrecisionAsDisplayed(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "PrecisionAsDisplayed"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPrecisionAsDisplayed(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(405) //= 0x195. The runtime will prefer the VTID if present
  @VTID(84)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setPrecisionAsDisplayed(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "PrecisionAsDisplayed"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(405) //= 0x195. The runtime will prefer the VTID if present
  @VTID(84)
  void setPrecisionAsDisplayed(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(85)
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter enableChanges is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void printPreview();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(enableChanges, 1033);
   * </code>
   * </p>
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges);

  /**
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(86)
  void printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter structure is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _Protect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter structure is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, structure, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param structure Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object structure);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param structure Optional parameter. Default value is com4j.Variant.getMissing()
   * @param windows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(87)
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object structure,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object windows);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ProtectSharing();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(filename, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(filename, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(filename, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ProtectSharing(filename, password, writeResPassword, readOnlyRecommended, createBackup, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sharingPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1450) //= 0x5aa. The runtime will prefer the VTID if present
  @VTID(88)
  void _ProtectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sharingPassword);


  /**
   * <p>
   * Getter method for the COM property "ProtectStructure"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(588) //= 0x24c. The runtime will prefer the VTID if present
  @VTID(89)
  boolean getProtectStructure();


  /**
   * <p>
   * Getter method for the COM property "ProtectWindows"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(295) //= 0x127. The runtime will prefer the VTID if present
  @VTID(90)
  boolean getProtectWindows();


  /**
   * <p>
   * Getter method for the COM property "ReadOnly"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getReadOnly(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(296) //= 0x128. The runtime will prefer the VTID if present
  @VTID(91)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getReadOnly();

  /**
   * <p>
   * Getter method for the COM property "ReadOnly"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(296) //= 0x128. The runtime will prefer the VTID if present
  @VTID(91)
  boolean getReadOnly(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "_ReadOnlyRecommended"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_ReadOnlyRecommended(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(297) //= 0x129. The runtime will prefer the VTID if present
  @VTID(92)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean get_ReadOnlyRecommended();

  /**
   * <p>
   * Getter method for the COM property "_ReadOnlyRecommended"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(297) //= 0x129. The runtime will prefer the VTID if present
  @VTID(92)
  boolean get_ReadOnlyRecommended(
    @LCID int lcid);


  /**
   */

  @DISPID(1452) //= 0x5ac. The runtime will prefer the VTID if present
  @VTID(93)
  void refreshAll();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * reply(1033);
   * </code>
   * </p>
   */

  @DISPID(977) //= 0x3d1. The runtime will prefer the VTID if present
  @VTID(94)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void reply();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(977) //= 0x3d1. The runtime will prefer the VTID if present
  @VTID(94)
  void reply(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replyAll(1033);
   * </code>
   * </p>
   */

  @DISPID(978) //= 0x3d2. The runtime will prefer the VTID if present
  @VTID(95)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void replyAll();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(978) //= 0x3d2. The runtime will prefer the VTID if present
  @VTID(95)
  void replyAll(
    @LCID int lcid);


  /**
   * @param index Mandatory int parameter.
   */

  @DISPID(1453) //= 0x5ad. The runtime will prefer the VTID if present
  @VTID(96)
  void removeUser(
    int index);


  /**
   * <p>
   * Getter method for the COM property "RevisionNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRevisionNumber(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1172) //= 0x494. The runtime will prefer the VTID if present
  @VTID(97)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getRevisionNumber();

  /**
   * <p>
   * Getter method for the COM property "RevisionNumber"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1172) //= 0x494. The runtime will prefer the VTID if present
  @VTID(97)
  int getRevisionNumber(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * route(1033);
   * </code>
   * </p>
   */

  @DISPID(946) //= 0x3b2. The runtime will prefer the VTID if present
  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void route();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(946) //= 0x3b2. The runtime will prefer the VTID if present
  @VTID(98)
  void route(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Routed"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRouted(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(951) //= 0x3b7. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getRouted();

  /**
   * <p>
   * Getter method for the COM property "Routed"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(951) //= 0x3b7. The runtime will prefer the VTID if present
  @VTID(99)
  boolean getRouted(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "RoutingSlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.RoutingSlip
   */

  @DISPID(949) //= 0x3b5. The runtime will prefer the VTID if present
  @VTID(100)
  com.exceljava.com4j.excel.RoutingSlip getRoutingSlip();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * runAutoMacros(which, 1033);
   * </code>
   * </p>
   * @param which Mandatory com.exceljava.com4j.excel.XlRunAutoMacro parameter.
   */

  @DISPID(634) //= 0x27a. The runtime will prefer the VTID if present
  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void runAutoMacros(
    com.exceljava.com4j.excel.XlRunAutoMacro which);

  /**
   * @param which Mandatory com.exceljava.com4j.excel.XlRunAutoMacro parameter.
   * @param lcid Mandatory int parameter.
   */

  @DISPID(634) //= 0x27a. The runtime will prefer the VTID if present
  @VTID(101)
  void runAutoMacros(
    com.exceljava.com4j.excel.XlRunAutoMacro which,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * save(1033);
   * </code>
   * </p>
   */

  @DISPID(283) //= 0x11b. The runtime will prefer the VTID if present
  @VTID(102)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void save();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(283) //= 0x11b. The runtime will prefer the VTID if present
  @VTID(102)
  void save(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11}, javaType = {com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(103)
  void __SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveCopyAs(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(175) //= 0xaf. The runtime will prefer the VTID if present
  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void saveCopyAs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveCopyAs(filename, 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(175) //= 0xaf. The runtime will prefer the VTID if present
  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void saveCopyAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(175) //= 0xaf. The runtime will prefer the VTID if present
  @VTID(104)
  void saveCopyAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Saved"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSaved(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(298) //= 0x12a. The runtime will prefer the VTID if present
  @VTID(105)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getSaved();

  /**
   * <p>
   * Getter method for the COM property "Saved"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(298) //= 0x12a. The runtime will prefer the VTID if present
  @VTID(105)
  boolean getSaved(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Saved"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSaved(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(298) //= 0x12a. The runtime will prefer the VTID if present
  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setSaved(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "Saved"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(298) //= 0x12a. The runtime will prefer the VTID if present
  @VTID(106)
  void setSaved(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveLinkValues"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSaveLinkValues(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(406) //= 0x196. The runtime will prefer the VTID if present
  @VTID(107)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getSaveLinkValues();

  /**
   * <p>
   * Getter method for the COM property "SaveLinkValues"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(406) //= 0x196. The runtime will prefer the VTID if present
  @VTID(107)
  boolean getSaveLinkValues(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "SaveLinkValues"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSaveLinkValues(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(406) //= 0x196. The runtime will prefer the VTID if present
  @VTID(108)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setSaveLinkValues(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "SaveLinkValues"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(406) //= 0x196. The runtime will prefer the VTID if present
  @VTID(108)
  void setSaveLinkValues(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subject is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter returnReceipt is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMail(recipients, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param recipients Mandatory java.lang.Object parameter.
   */

  @DISPID(947) //= 0x3b3. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void sendMail(
    @MarshalAs(NativeType.VARIANT) java.lang.Object recipients);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter returnReceipt is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMail(recipients, subject, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param recipients Mandatory java.lang.Object parameter.
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(947) //= 0x3b3. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void sendMail(
    @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMail(recipients, subject, returnReceipt, 1033);
   * </code>
   * </p>
   * @param recipients Mandatory java.lang.Object parameter.
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   * @param returnReceipt Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(947) //= 0x3b3. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void sendMail(
    @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object returnReceipt);

  /**
   * @param recipients Mandatory java.lang.Object parameter.
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   * @param returnReceipt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(947) //= 0x3b3. The runtime will prefer the VTID if present
  @VTID(109)
  void sendMail(
    @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object returnReceipt,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlPriority parameter priority is set to -4143</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMailer(com4j.Variant.getMissing(), -4143, 1033);
   * </code>
   * </p>
   */

  @DISPID(980) //= 0x3d4. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPriority.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "-4143", "1033"})
  void sendMailer();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPriority parameter priority is set to -4143</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMailer(fileFormat, -4143, 1033);
   * </code>
   * </p>
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(980) //= 0x3d4. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {com.exceljava.com4j.excel.XlPriority.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-4143", "1033"})
  void sendMailer(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendMailer(fileFormat, priority, 1033);
   * </code>
   * </p>
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is -4143
   */

  @DISPID(980) //= 0x3d4. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void sendMailer(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @DefaultValue("-4143") com.exceljava.com4j.excel.XlPriority priority);

  /**
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param priority Optional parameter. Default value is -4143
   * @param lcid Mandatory int parameter.
   */

  @DISPID(980) //= 0x3d4. The runtime will prefer the VTID if present
  @VTID(110)
  void sendMailer(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @DefaultValue("-4143") com.exceljava.com4j.excel.XlPriority priority,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter procedure is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setLinkOnData(name, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   */

  @DISPID(809) //= 0x329. The runtime will prefer the VTID if present
  @VTID(111)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void setLinkOnData(
    java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setLinkOnData(name, procedure, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param procedure Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(809) //= 0x329. The runtime will prefer the VTID if present
  @VTID(111)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setLinkOnData(
    java.lang.String name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object procedure);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param procedure Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(809) //= 0x329. The runtime will prefer the VTID if present
  @VTID(111)
  void setLinkOnData(
    java.lang.String name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object procedure,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Sheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(485) //= 0x1e5. The runtime will prefer the VTID if present
  @VTID(112)
  com.exceljava.com4j.excel.Sheets getSheets();


  @VTID(112)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getSheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "ShowConflictHistory"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getShowConflictHistory(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1171) //= 0x493. The runtime will prefer the VTID if present
  @VTID(113)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getShowConflictHistory();

  /**
   * <p>
   * Getter method for the COM property "ShowConflictHistory"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1171) //= 0x493. The runtime will prefer the VTID if present
  @VTID(113)
  boolean getShowConflictHistory(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "ShowConflictHistory"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setShowConflictHistory(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1171) //= 0x493. The runtime will prefer the VTID if present
  @VTID(114)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setShowConflictHistory(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "ShowConflictHistory"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1171) //= 0x493. The runtime will prefer the VTID if present
  @VTID(114)
  void setShowConflictHistory(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Styles"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Styles
   */

  @DISPID(493) //= 0x1ed. The runtime will prefer the VTID if present
  @VTID(115)
  com.exceljava.com4j.excel.Styles getStyles();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSubject(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(953) //= 0x3b9. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getSubject();

  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(953) //= 0x3b9. The runtime will prefer the VTID if present
  @VTID(116)
  java.lang.String getSubject(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSubject(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(953) //= 0x3b9. The runtime will prefer the VTID if present
  @VTID(117)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setSubject(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(953) //= 0x3b9. The runtime will prefer the VTID if present
  @VTID(117)
  void setSubject(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getTitle(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(199) //= 0xc7. The runtime will prefer the VTID if present
  @VTID(118)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getTitle();

  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(199) //= 0xc7. The runtime will prefer the VTID if present
  @VTID(118)
  java.lang.String getTitle(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setTitle(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(199) //= 0xc7. The runtime will prefer the VTID if present
  @VTID(119)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setTitle(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(199) //= 0xc7. The runtime will prefer the VTID if present
  @VTID(119)
  void setTitle(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotect(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(120)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void unprotect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotect(password, 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(120)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void unprotect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(120)
  void unprotect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotectSharing(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1455) //= 0x5af. The runtime will prefer the VTID if present
  @VTID(121)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void unprotectSharing();

  /**
   * @param sharingPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1455) //= 0x5af. The runtime will prefer the VTID if present
  @VTID(121)
  void unprotectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sharingPassword);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * updateFromFile(1033);
   * </code>
   * </p>
   */

  @DISPID(995) //= 0x3e3. The runtime will prefer the VTID if present
  @VTID(122)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void updateFromFile();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(995) //= 0x3e3. The runtime will prefer the VTID if present
  @VTID(122)
  void updateFromFile(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * updateLink(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(804) //= 0x324. The runtime will prefer the VTID if present
  @VTID(123)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void updateLink();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * updateLink(name, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(804) //= 0x324. The runtime will prefer the VTID if present
  @VTID(123)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void updateLink(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * updateLink(name, type, 1033);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(804) //= 0x324. The runtime will prefer the VTID if present
  @VTID(123)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void updateLink(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(804) //= 0x324. The runtime will prefer the VTID if present
  @VTID(123)
  void updateLink(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "UpdateRemoteReferences"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getUpdateRemoteReferences(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(411) //= 0x19b. The runtime will prefer the VTID if present
  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getUpdateRemoteReferences();

  /**
   * <p>
   * Getter method for the COM property "UpdateRemoteReferences"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(411) //= 0x19b. The runtime will prefer the VTID if present
  @VTID(124)
  boolean getUpdateRemoteReferences(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "UpdateRemoteReferences"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setUpdateRemoteReferences(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(411) //= 0x19b. The runtime will prefer the VTID if present
  @VTID(125)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setUpdateRemoteReferences(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "UpdateRemoteReferences"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(411) //= 0x19b. The runtime will prefer the VTID if present
  @VTID(125)
  void setUpdateRemoteReferences(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UserControl"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1210) //= 0x4ba. The runtime will prefer the VTID if present
  @VTID(126)
  boolean getUserControl();


  /**
   * <p>
   * Setter method for the COM property "UserControl"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1210) //= 0x4ba. The runtime will prefer the VTID if present
  @VTID(127)
  void setUserControl(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UserStatus"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getUserStatus(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1173) //= 0x495. The runtime will prefer the VTID if present
  @VTID(128)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getUserStatus();

  /**
   * <p>
   * Getter method for the COM property "UserStatus"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1173) //= 0x495. The runtime will prefer the VTID if present
  @VTID(128)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getUserStatus(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CustomViews"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CustomViews
   */

  @DISPID(1456) //= 0x5b0. The runtime will prefer the VTID if present
  @VTID(129)
  com.exceljava.com4j.excel.CustomViews getCustomViews();


  /**
   * <p>
   * Getter method for the COM property "Windows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Windows
   */

  @DISPID(430) //= 0x1ae. The runtime will prefer the VTID if present
  @VTID(130)
  com.exceljava.com4j.excel.Windows getWindows();


  /**
   * <p>
   * Getter method for the COM property "Worksheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(494) //= 0x1ee. The runtime will prefer the VTID if present
  @VTID(131)
  com.exceljava.com4j.excel.Sheets getWorksheets();


  @VTID(131)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getWorksheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "WriteReserved"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWriteReserved(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(299) //= 0x12b. The runtime will prefer the VTID if present
  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getWriteReserved();

  /**
   * <p>
   * Getter method for the COM property "WriteReserved"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(299) //= 0x12b. The runtime will prefer the VTID if present
  @VTID(132)
  boolean getWriteReserved(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "WriteReservedBy"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWriteReservedBy(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(300) //= 0x12c. The runtime will prefer the VTID if present
  @VTID(133)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getWriteReservedBy();

  /**
   * <p>
   * Getter method for the COM property "WriteReservedBy"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(300) //= 0x12c. The runtime will prefer the VTID if present
  @VTID(133)
  java.lang.String getWriteReservedBy(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Excel4IntlMacroSheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(581) //= 0x245. The runtime will prefer the VTID if present
  @VTID(134)
  com.exceljava.com4j.excel.Sheets getExcel4IntlMacroSheets();


  @VTID(134)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getExcel4IntlMacroSheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Excel4MacroSheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @DISPID(579) //= 0x243. The runtime will prefer the VTID if present
  @VTID(135)
  com.exceljava.com4j.excel.Sheets getExcel4MacroSheets();


  @VTID(135)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getExcel4MacroSheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "TemplateRemoveExtData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1457) //= 0x5b1. The runtime will prefer the VTID if present
  @VTID(136)
  boolean getTemplateRemoveExtData();


  /**
   * <p>
   * Setter method for the COM property "TemplateRemoveExtData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1457) //= 0x5b1. The runtime will prefer the VTID if present
  @VTID(137)
  void setTemplateRemoveExtData(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter when is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * highlightChangesOptions(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1458) //= 0x5b2. The runtime will prefer the VTID if present
  @VTID(138)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void highlightChangesOptions();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * highlightChangesOptions(when, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1458) //= 0x5b2. The runtime will prefer the VTID if present
  @VTID(138)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void highlightChangesOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * highlightChangesOptions(when, who, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1458) //= 0x5b2. The runtime will prefer the VTID if present
  @VTID(138)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void highlightChangesOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who);

  /**
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   * @param where Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1458) //= 0x5b2. The runtime will prefer the VTID if present
  @VTID(138)
  void highlightChangesOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object where);


  /**
   * <p>
   * Getter method for the COM property "HighlightChangesOnScreen"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1461) //= 0x5b5. The runtime will prefer the VTID if present
  @VTID(139)
  boolean getHighlightChangesOnScreen();


  /**
   * <p>
   * Setter method for the COM property "HighlightChangesOnScreen"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1461) //= 0x5b5. The runtime will prefer the VTID if present
  @VTID(140)
  void setHighlightChangesOnScreen(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "KeepChangeHistory"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1462) //= 0x5b6. The runtime will prefer the VTID if present
  @VTID(141)
  boolean getKeepChangeHistory();


  /**
   * <p>
   * Setter method for the COM property "KeepChangeHistory"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1462) //= 0x5b6. The runtime will prefer the VTID if present
  @VTID(142)
  void setKeepChangeHistory(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListChangesOnNewSheet"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1463) //= 0x5b7. The runtime will prefer the VTID if present
  @VTID(143)
  boolean getListChangesOnNewSheet();


  /**
   * <p>
   * Setter method for the COM property "ListChangesOnNewSheet"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1463) //= 0x5b7. The runtime will prefer the VTID if present
  @VTID(144)
  void setListChangesOnNewSheet(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * purgeChangeHistoryNow(days, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param days Mandatory int parameter.
   */

  @DISPID(1464) //= 0x5b8. The runtime will prefer the VTID if present
  @VTID(145)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void purgeChangeHistoryNow(
    int days);

  /**
   * @param days Mandatory int parameter.
   * @param sharingPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1464) //= 0x5b8. The runtime will prefer the VTID if present
  @VTID(145)
  void purgeChangeHistoryNow(
    int days,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sharingPassword);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter when is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * acceptAllChanges(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1466) //= 0x5ba. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void acceptAllChanges();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * acceptAllChanges(when, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1466) //= 0x5ba. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void acceptAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * acceptAllChanges(when, who, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1466) //= 0x5ba. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void acceptAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who);

  /**
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   * @param where Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1466) //= 0x5ba. The runtime will prefer the VTID if present
  @VTID(146)
  void acceptAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object where);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter when is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * rejectAllChanges(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1467) //= 0x5bb. The runtime will prefer the VTID if present
  @VTID(147)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void rejectAllChanges();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter who is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * rejectAllChanges(when, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1467) //= 0x5bb. The runtime will prefer the VTID if present
  @VTID(147)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void rejectAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter where is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * rejectAllChanges(when, who, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1467) //= 0x5bb. The runtime will prefer the VTID if present
  @VTID(147)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void rejectAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who);

  /**
   * @param when Optional parameter. Default value is com4j.Variant.getMissing()
   * @param who Optional parameter. Default value is com4j.Variant.getMissing()
   * @param where Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1467) //= 0x5bb. The runtime will prefer the VTID if present
  @VTID(147)
  void rejectAllChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object when,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object who,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object where);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15, 16}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection, 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param connection Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, optParamIndex = {16}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object connection);

  /**
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param connection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(148)
  void pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object connection,
    @LCID int lcid);


  /**
   */

  @DISPID(1468) //= 0x5bc. The runtime will prefer the VTID if present
  @VTID(149)
  void resetColors();


  /**
   * <p>
   * Getter method for the COM property "VBProject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(1469) //= 0x5bd. The runtime will prefer the VTID if present
  @VTID(150)
  com.exceljava.com4j.vbide._VBProject getVBProject();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subAddress is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newWindow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addHistory is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void followHyperlink(
    java.lang.String address);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newWindow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addHistory is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, subAddress, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addHistory is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, subAddress, newWindow, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newWindow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, subAddress, newWindow, addHistory, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newWindow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addHistory);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, subAddress, newWindow, addHistory, extraInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newWindow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addHistory,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object extraInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * followHyperlink(address, subAddress, newWindow, addHistory, extraInfo, method, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param method Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newWindow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addHistory,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object extraInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object method);

  /**
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param method Optional parameter. Default value is com4j.Variant.getMissing()
   * @param headerInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1470) //= 0x5be. The runtime will prefer the VTID if present
  @VTID(151)
  void followHyperlink(
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newWindow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addHistory,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object extraInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object method,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object headerInfo);


  /**
   */

  @DISPID(1476) //= 0x5c4. The runtime will prefer the VTID if present
  @VTID(152)
  void addToFavorites();


  /**
   * <p>
   * Getter method for the COM property "IsInplace"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1769) //= 0x6e9. The runtime will prefer the VTID if present
  @VTID(153)
  boolean getIsInplace();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(154)
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @LCID int lcid);


  /**
   */

  @DISPID(1818) //= 0x71a. The runtime will prefer the VTID if present
  @VTID(155)
  void webPagePreview();


  /**
   * <p>
   * Getter method for the COM property "PublishObjects"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishObjects
   */

  @DISPID(1819) //= 0x71b. The runtime will prefer the VTID if present
  @VTID(156)
  com.exceljava.com4j.excel.PublishObjects getPublishObjects();


  /**
   * <p>
   * Getter method for the COM property "WebOptions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WebOptions
   */

  @DISPID(1820) //= 0x71c. The runtime will prefer the VTID if present
  @VTID(157)
  com.exceljava.com4j.excel.WebOptions getWebOptions();


  /**
   * @param encoding Mandatory com.exceljava.com4j.office.MsoEncoding parameter.
   */

  @DISPID(1821) //= 0x71d. The runtime will prefer the VTID if present
  @VTID(158)
  void reloadAs(
    com.exceljava.com4j.office.MsoEncoding encoding);


  /**
   * <p>
   * Getter method for the COM property "HTMLProject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.HTMLProject
   */

  @DISPID(1823) //= 0x71f. The runtime will prefer the VTID if present
  @VTID(159)
  com.exceljava.com4j.office.HTMLProject getHTMLProject();


  /**
   * <p>
   * Getter method for the COM property "EnvelopeVisible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1824) //= 0x720. The runtime will prefer the VTID if present
  @VTID(160)
  boolean getEnvelopeVisible();


  /**
   * <p>
   * Setter method for the COM property "EnvelopeVisible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1824) //= 0x720. The runtime will prefer the VTID if present
  @VTID(161)
  void setEnvelopeVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CalculationVersion"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1806) //= 0x70e. The runtime will prefer the VTID if present
  @VTID(162)
  int getCalculationVersion();


  /**
   * @param calcid Mandatory int parameter.
   */

  @DISPID(2044) //= 0x7fc. The runtime will prefer the VTID if present
  @VTID(163)
  void dummy17(
    int calcid);


  /**
   * @param s Mandatory java.lang.String parameter.
   */

  @DISPID(1826) //= 0x722. The runtime will prefer the VTID if present
  @VTID(164)
  void sblt(
    java.lang.String s);


  /**
   * <p>
   * Getter method for the COM property "VBASigned"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1828) //= 0x724. The runtime will prefer the VTID if present
  @VTID(165)
  boolean getVBASigned();


  /**
   * <p>
   * Getter method for the COM property "ShowPivotTableFieldList"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2046) //= 0x7fe. The runtime will prefer the VTID if present
  @VTID(166)
  boolean getShowPivotTableFieldList();


  /**
   * <p>
   * Setter method for the COM property "ShowPivotTableFieldList"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2046) //= 0x7fe. The runtime will prefer the VTID if present
  @VTID(167)
  void setShowPivotTableFieldList(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UpdateLinks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlUpdateLinks
   */

  @DISPID(864) //= 0x360. The runtime will prefer the VTID if present
  @VTID(168)
  com.exceljava.com4j.excel.XlUpdateLinks getUpdateLinks();


  /**
   * <p>
   * Setter method for the COM property "UpdateLinks"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlUpdateLinks parameter.
   */

  @DISPID(864) //= 0x360. The runtime will prefer the VTID if present
  @VTID(169)
  void setUpdateLinks(
    com.exceljava.com4j.excel.XlUpdateLinks rhs);


  /**
   * @param name Mandatory java.lang.String parameter.
   * @param type Mandatory com.exceljava.com4j.excel.XlLinkType parameter.
   */

  @DISPID(2047) //= 0x7ff. The runtime will prefer the VTID if present
  @VTID(170)
  void breakLink(
    java.lang.String name,
    com.exceljava.com4j.excel.XlLinkType type);


  /**
   */

  @DISPID(2048) //= 0x800. The runtime will prefer the VTID if present
  @VTID(171)
  void dummy16();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12}, javaType = {com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local, 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(172)
  void _SaveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "EnableAutoRecover"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2049) //= 0x801. The runtime will prefer the VTID if present
  @VTID(173)
  boolean getEnableAutoRecover();


  /**
   * <p>
   * Setter method for the COM property "EnableAutoRecover"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2049) //= 0x801. The runtime will prefer the VTID if present
  @VTID(174)
  void setEnableAutoRecover(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RemovePersonalInformation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2050) //= 0x802. The runtime will prefer the VTID if present
  @VTID(175)
  boolean getRemovePersonalInformation();


  /**
   * <p>
   * Setter method for the COM property "RemovePersonalInformation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2050) //= 0x802. The runtime will prefer the VTID if present
  @VTID(176)
  void setRemovePersonalInformation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FullNameURLEncoded"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFullNameURLEncoded(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1927) //= 0x787. The runtime will prefer the VTID if present
  @VTID(177)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getFullNameURLEncoded();

  /**
   * <p>
   * Getter method for the COM property "FullNameURLEncoded"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1927) //= 0x787. The runtime will prefer the VTID if present
  @VTID(177)
  java.lang.String getFullNameURLEncoded(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveChanges is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comments is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkIn(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2051) //= 0x803. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void checkIn();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comments is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkIn(saveChanges, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2051) //= 0x803. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void checkIn(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkIn(saveChanges, comments, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comments Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2051) //= 0x803. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void checkIn(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comments);

  /**
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comments Optional parameter. Default value is com4j.Variant.getMissing()
   * @param makePublic Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2051) //= 0x803. The runtime will prefer the VTID if present
  @VTID(178)
  void checkIn(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comments,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object makePublic);


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(2053) //= 0x805. The runtime will prefer the VTID if present
  @VTID(179)
  boolean canCheckIn();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter recipients is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subject is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeAttachment is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendForReview(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2054) //= 0x806. The runtime will prefer the VTID if present
  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void sendForReview();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subject is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeAttachment is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendForReview(recipients, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2054) //= 0x806. The runtime will prefer the VTID if present
  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void sendForReview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeAttachment is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendForReview(recipients, subject, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2054) //= 0x806. The runtime will prefer the VTID if present
  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void sendForReview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter includeAttachment is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendForReview(recipients, subject, showMessage, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showMessage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2054) //= 0x806. The runtime will prefer the VTID if present
  @VTID(180)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void sendForReview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showMessage);

  /**
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showMessage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeAttachment Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2054) //= 0x806. The runtime will prefer the VTID if present
  @VTID(180)
  void sendForReview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showMessage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeAttachment);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replyWithChanges(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2057) //= 0x809. The runtime will prefer the VTID if present
  @VTID(181)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void replyWithChanges();

  /**
   * @param showMessage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2057) //= 0x809. The runtime will prefer the VTID if present
  @VTID(181)
  void replyWithChanges(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showMessage);


  /**
   */

  @DISPID(2058) //= 0x80a. The runtime will prefer the VTID if present
  @VTID(182)
  void endReview();


  /**
   * <p>
   * Getter method for the COM property "Password"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(429) //= 0x1ad. The runtime will prefer the VTID if present
  @VTID(183)
  java.lang.String getPassword();


  /**
   * <p>
   * Setter method for the COM property "Password"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(429) //= 0x1ad. The runtime will prefer the VTID if present
  @VTID(184)
  void setPassword(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "WritePassword"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1128) //= 0x468. The runtime will prefer the VTID if present
  @VTID(185)
  java.lang.String getWritePassword();


  /**
   * <p>
   * Setter method for the COM property "WritePassword"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1128) //= 0x468. The runtime will prefer the VTID if present
  @VTID(186)
  void setWritePassword(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "PasswordEncryptionProvider"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2059) //= 0x80b. The runtime will prefer the VTID if present
  @VTID(187)
  java.lang.String getPasswordEncryptionProvider();


  /**
   * <p>
   * Getter method for the COM property "PasswordEncryptionAlgorithm"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2060) //= 0x80c. The runtime will prefer the VTID if present
  @VTID(188)
  java.lang.String getPasswordEncryptionAlgorithm();


  /**
   * <p>
   * Getter method for the COM property "PasswordEncryptionKeyLength"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2061) //= 0x80d. The runtime will prefer the VTID if present
  @VTID(189)
  int getPasswordEncryptionKeyLength();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter passwordEncryptionProvider is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionAlgorithm is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionKeyLength is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionFileProperties is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPasswordEncryptionOptions(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2062) //= 0x80e. The runtime will prefer the VTID if present
  @VTID(190)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void setPasswordEncryptionOptions();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter passwordEncryptionAlgorithm is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionKeyLength is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionFileProperties is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPasswordEncryptionOptions(passwordEncryptionProvider, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param passwordEncryptionProvider Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2062) //= 0x80e. The runtime will prefer the VTID if present
  @VTID(190)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void setPasswordEncryptionOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionProvider);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter passwordEncryptionKeyLength is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter passwordEncryptionFileProperties is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPasswordEncryptionOptions(passwordEncryptionProvider, passwordEncryptionAlgorithm, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param passwordEncryptionProvider Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionAlgorithm Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2062) //= 0x80e. The runtime will prefer the VTID if present
  @VTID(190)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void setPasswordEncryptionOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionProvider,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionAlgorithm);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter passwordEncryptionFileProperties is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPasswordEncryptionOptions(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param passwordEncryptionProvider Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionAlgorithm Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionKeyLength Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2062) //= 0x80e. The runtime will prefer the VTID if present
  @VTID(190)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setPasswordEncryptionOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionProvider,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionAlgorithm,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionKeyLength);

  /**
   * @param passwordEncryptionProvider Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionAlgorithm Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionKeyLength Optional parameter. Default value is com4j.Variant.getMissing()
   * @param passwordEncryptionFileProperties Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2062) //= 0x80e. The runtime will prefer the VTID if present
  @VTID(190)
  void setPasswordEncryptionOptions(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionProvider,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionAlgorithm,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionKeyLength,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object passwordEncryptionFileProperties);


  /**
   * <p>
   * Getter method for the COM property "PasswordEncryptionFileProperties"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2063) //= 0x80f. The runtime will prefer the VTID if present
  @VTID(191)
  boolean getPasswordEncryptionFileProperties();


  /**
   * <p>
   * Getter method for the COM property "ReadOnlyRecommended"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2005) //= 0x7d5. The runtime will prefer the VTID if present
  @VTID(192)
  boolean getReadOnlyRecommended();


  /**
   * <p>
   * Setter method for the COM property "ReadOnlyRecommended"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2005) //= 0x7d5. The runtime will prefer the VTID if present
  @VTID(193)
  void setReadOnlyRecommended(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter structure is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(194)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void protect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter structure is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(194)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter windows is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, structure, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param structure Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(194)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object structure);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param structure Optional parameter. Default value is com4j.Variant.getMissing()
   * @param windows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(194)
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object structure,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object windows);


  /**
   * <p>
   * Getter method for the COM property "SmartTagOptions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SmartTagOptions
   */

  @DISPID(2064) //= 0x810. The runtime will prefer the VTID if present
  @VTID(195)
  com.exceljava.com4j.excel.SmartTagOptions getSmartTagOptions();


  /**
   */

  @DISPID(2065) //= 0x811. The runtime will prefer the VTID if present
  @VTID(196)
  void recheckSmartTags();


  /**
   * <p>
   * Getter method for the COM property "Permission"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Permission
   */

  @DISPID(2264) //= 0x8d8. The runtime will prefer the VTID if present
  @VTID(197)
  com.exceljava.com4j.office.Permission getPermission();


  @VTID(197)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.Permission.class})
  com.exceljava.com4j.office.UserPermission getPermission(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "SharedWorkspace"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspace
   */

  @DISPID(2265) //= 0x8d9. The runtime will prefer the VTID if present
  @VTID(198)
  com.exceljava.com4j.office.SharedWorkspace getSharedWorkspace();


  /**
   * <p>
   * Getter method for the COM property "Sync"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Sync
   */

  @DISPID(2266) //= 0x8da. The runtime will prefer the VTID if present
  @VTID(199)
  com.exceljava.com4j.office.Sync getSync();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter recipients is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subject is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendFaxOverInternet(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2267) //= 0x8db. The runtime will prefer the VTID if present
  @VTID(200)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void sendFaxOverInternet();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subject is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendFaxOverInternet(recipients, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2267) //= 0x8db. The runtime will prefer the VTID if present
  @VTID(200)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void sendFaxOverInternet(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showMessage is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sendFaxOverInternet(recipients, subject, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2267) //= 0x8db. The runtime will prefer the VTID if present
  @VTID(200)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void sendFaxOverInternet(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject);

  /**
   * @param recipients Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subject Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showMessage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2267) //= 0x8db. The runtime will prefer the VTID if present
  @VTID(200)
  void sendFaxOverInternet(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recipients,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subject,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showMessage);


  /**
   * <p>
   * Getter method for the COM property "XmlNamespaces"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlNamespaces
   */

  @DISPID(2268) //= 0x8dc. The runtime will prefer the VTID if present
  @VTID(201)
  com.exceljava.com4j.excel.XmlNamespaces getXmlNamespaces();


  /**
   * <p>
   * Getter method for the COM property "XmlMaps"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMaps
   */

  @DISPID(2269) //= 0x8dd. The runtime will prefer the VTID if present
  @VTID(202)
  com.exceljava.com4j.excel.XmlMaps getXmlMaps();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlImport(url, importMap, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param url Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2270) //= 0x8de. The runtime will prefer the VTID if present
  @VTID(203)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImport(
    java.lang.String url,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlImport(url, importMap, overwrite, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param url Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2270) //= 0x8de. The runtime will prefer the VTID if present
  @VTID(203)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImport(
    java.lang.String url,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwrite);

  /**
   * @param url Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2270) //= 0x8de. The runtime will prefer the VTID if present
  @VTID(203)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImport(
    java.lang.String url,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwrite,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


  /**
   * <p>
   * Getter method for the COM property "SmartDocument"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartDocument
   */

  @DISPID(2273) //= 0x8e1. The runtime will prefer the VTID if present
  @VTID(204)
  com.exceljava.com4j.office.SmartDocument getSmartDocument();


  /**
   * <p>
   * Getter method for the COM property "DocumentLibraryVersions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentLibraryVersions
   */

  @DISPID(2274) //= 0x8e2. The runtime will prefer the VTID if present
  @VTID(205)
  com.exceljava.com4j.office.DocumentLibraryVersions getDocumentLibraryVersions();


  @VTID(205)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.DocumentLibraryVersions.class})
  com.exceljava.com4j.office.DocumentLibraryVersion getDocumentLibraryVersions(
    int lIndex);

  /**
   * <p>
   * Getter method for the COM property "InactiveListBorderVisible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2275) //= 0x8e3. The runtime will prefer the VTID if present
  @VTID(206)
  boolean getInactiveListBorderVisible();


  /**
   * <p>
   * Setter method for the COM property "InactiveListBorderVisible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2275) //= 0x8e3. The runtime will prefer the VTID if present
  @VTID(207)
  void setInactiveListBorderVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayInkComments"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2276) //= 0x8e4. The runtime will prefer the VTID if present
  @VTID(208)
  boolean getDisplayInkComments();


  /**
   * <p>
   * Setter method for the COM property "DisplayInkComments"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2276) //= 0x8e4. The runtime will prefer the VTID if present
  @VTID(209)
  void setDisplayInkComments(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlImportXml(data, importMap, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param data Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2277) //= 0x8e5. The runtime will prefer the VTID if present
  @VTID(210)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImportXml(
    java.lang.String data,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlImportXml(data, importMap, overwrite, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param data Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2277) //= 0x8e5. The runtime will prefer the VTID if present
  @VTID(210)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImportXml(
    java.lang.String data,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwrite);

  /**
   * @param data Mandatory java.lang.String parameter.
   * @param importMap Mandatory Holder<com.exceljava.com4j.excel.XmlMap> parameter.
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.XlXmlImportResult
   */

  @DISPID(2277) //= 0x8e5. The runtime will prefer the VTID if present
  @VTID(210)
  com.exceljava.com4j.excel.XlXmlImportResult xmlImportXml(
    java.lang.String data,
    Holder<com.exceljava.com4j.excel.XmlMap> importMap,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwrite,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   */

  @DISPID(2278) //= 0x8e6. The runtime will prefer the VTID if present
  @VTID(211)
  void saveAsXMLData(
    java.lang.String filename,
    com.exceljava.com4j.excel.XmlMap map);


  /**
   */

  @DISPID(2279) //= 0x8e7. The runtime will prefer the VTID if present
  @VTID(212)
  void toggleFormsDesign();


  /**
   * <p>
   * Getter method for the COM property "ContentTypeProperties"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MetaProperties
   */

  @DISPID(2512) //= 0x9d0. The runtime will prefer the VTID if present
  @VTID(213)
  com.exceljava.com4j.office.MetaProperties getContentTypeProperties();


  @VTID(213)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.MetaProperties.class})
  com.exceljava.com4j.office.MetaProperty getContentTypeProperties(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Connections"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Connections
   */

  @DISPID(2513) //= 0x9d1. The runtime will prefer the VTID if present
  @VTID(214)
  com.exceljava.com4j.excel.Connections getConnections();


  /**
   * @param removeDocInfoType Mandatory com.exceljava.com4j.excel.XlRemoveDocInfoType parameter.
   */

  @DISPID(2514) //= 0x9d2. The runtime will prefer the VTID if present
  @VTID(215)
  void removeDocumentInformation(
    com.exceljava.com4j.excel.XlRemoveDocInfoType removeDocInfoType);


  /**
   * <p>
   * Getter method for the COM property "Signatures"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SignatureSet
   */

  @DISPID(2516) //= 0x9d4. The runtime will prefer the VTID if present
  @VTID(216)
  com.exceljava.com4j.office.SignatureSet getSignatures();


  @VTID(216)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SignatureSet.class})
  com.exceljava.com4j.office.Signature getSignatures(
    int iSig);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveChanges is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comments is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter versionType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkInWithVersion(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2517) //= 0x9d5. The runtime will prefer the VTID if present
  @VTID(217)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void checkInWithVersion();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comments is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter versionType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkInWithVersion(saveChanges, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2517) //= 0x9d5. The runtime will prefer the VTID if present
  @VTID(217)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void checkInWithVersion(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter makePublic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter versionType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkInWithVersion(saveChanges, comments, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comments Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2517) //= 0x9d5. The runtime will prefer the VTID if present
  @VTID(217)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void checkInWithVersion(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comments);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter versionType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkInWithVersion(saveChanges, comments, makePublic, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comments Optional parameter. Default value is com4j.Variant.getMissing()
   * @param makePublic Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2517) //= 0x9d5. The runtime will prefer the VTID if present
  @VTID(217)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void checkInWithVersion(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comments,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object makePublic);

  /**
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comments Optional parameter. Default value is com4j.Variant.getMissing()
   * @param makePublic Optional parameter. Default value is com4j.Variant.getMissing()
   * @param versionType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2517) //= 0x9d5. The runtime will prefer the VTID if present
  @VTID(217)
  void checkInWithVersion(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comments,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object makePublic,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object versionType);


  /**
   * <p>
   * Getter method for the COM property "ServerPolicy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ServerPolicy
   */

  @DISPID(2519) //= 0x9d7. The runtime will prefer the VTID if present
  @VTID(218)
  com.exceljava.com4j.office.ServerPolicy getServerPolicy();


  @VTID(218)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.ServerPolicy.class})
  com.exceljava.com4j.office.PolicyItem getServerPolicy(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   */

  @DISPID(2520) //= 0x9d8. The runtime will prefer the VTID if present
  @VTID(219)
  void lockServerFile();


  /**
   * <p>
   * Getter method for the COM property "DocumentInspectors"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentInspectors
   */

  @DISPID(2521) //= 0x9d9. The runtime will prefer the VTID if present
  @VTID(220)
  com.exceljava.com4j.office.DocumentInspectors getDocumentInspectors();


  @VTID(220)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.DocumentInspectors.class})
  com.exceljava.com4j.office.DocumentInspector getDocumentInspectors(
    int index);

  /**
   * @return  Returns a value of type com.exceljava.com4j.office.WorkflowTasks
   */

  @DISPID(2522) //= 0x9da. The runtime will prefer the VTID if present
  @VTID(221)
  com.exceljava.com4j.office.WorkflowTasks getWorkflowTasks();


  @VTID(221)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.WorkflowTasks.class})
  com.exceljava.com4j.office.WorkflowTask getWorkflowTasks(
    int index);

  /**
   * @return  Returns a value of type com.exceljava.com4j.office.WorkflowTemplates
   */

  @DISPID(2523) //= 0x9db. The runtime will prefer the VTID if present
  @VTID(222)
  com.exceljava.com4j.office.WorkflowTemplates getWorkflowTemplates();


  @VTID(222)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.WorkflowTemplates.class})
  com.exceljava.com4j.office.WorkflowTemplate getWorkflowTemplates(
    int index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(223)
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ServerViewableItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ServerViewableItems
   */

  @DISPID(2524) //= 0x9dc. The runtime will prefer the VTID if present
  @VTID(224)
  com.exceljava.com4j.excel.ServerViewableItems getServerViewableItems();


  /**
   * <p>
   * Getter method for the COM property "TableStyles"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TableStyles
   */

  @DISPID(2525) //= 0x9dd. The runtime will prefer the VTID if present
  @VTID(225)
  com.exceljava.com4j.excel.TableStyles getTableStyles();


  /**
   * <p>
   * Getter method for the COM property "DefaultTableStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2526) //= 0x9de. The runtime will prefer the VTID if present
  @VTID(226)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDefaultTableStyle();


  /**
   * <p>
   * Setter method for the COM property "DefaultTableStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2526) //= 0x9de. The runtime will prefer the VTID if present
  @VTID(227)
  void setDefaultTableStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "DefaultPivotTableStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2527) //= 0x9df. The runtime will prefer the VTID if present
  @VTID(228)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDefaultPivotTableStyle();


  /**
   * <p>
   * Setter method for the COM property "DefaultPivotTableStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2527) //= 0x9df. The runtime will prefer the VTID if present
  @VTID(229)
  void setDefaultPivotTableStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CheckCompatibility"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2528) //= 0x9e0. The runtime will prefer the VTID if present
  @VTID(230)
  boolean getCheckCompatibility();


  /**
   * <p>
   * Setter method for the COM property "CheckCompatibility"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2528) //= 0x9e0. The runtime will prefer the VTID if present
  @VTID(231)
  void setCheckCompatibility(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasVBProject"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2529) //= 0x9e1. The runtime will prefer the VTID if present
  @VTID(232)
  boolean getHasVBProject();


  /**
   * <p>
   * Getter method for the COM property "CustomXMLParts"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLParts
   */

  @DISPID(2530) //= 0x9e2. The runtime will prefer the VTID if present
  @VTID(233)
  com.exceljava.com4j.office._CustomXMLParts getCustomXMLParts();


  /**
   * <p>
   * Getter method for the COM property "Final"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2531) //= 0x9e3. The runtime will prefer the VTID if present
  @VTID(234)
  boolean getFinal();


  /**
   * <p>
   * Setter method for the COM property "Final"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2531) //= 0x9e3. The runtime will prefer the VTID if present
  @VTID(235)
  void setFinal(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Research"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Research
   */

  @DISPID(2532) //= 0x9e4. The runtime will prefer the VTID if present
  @VTID(236)
  com.exceljava.com4j.excel.Research getResearch();


  /**
   * <p>
   * Getter method for the COM property "Theme"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.OfficeTheme
   */

  @DISPID(2533) //= 0x9e5. The runtime will prefer the VTID if present
  @VTID(237)
  com.exceljava.com4j.office.OfficeTheme getTheme();


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(2534) //= 0x9e6. The runtime will prefer the VTID if present
  @VTID(238)
  void applyTheme(
    java.lang.String filename);


  /**
   * <p>
   * Getter method for the COM property "Excel8CompatibilityMode"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2535) //= 0x9e7. The runtime will prefer the VTID if present
  @VTID(239)
  boolean getExcel8CompatibilityMode();


  /**
   * <p>
   * Getter method for the COM property "ConnectionsDisabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2536) //= 0x9e8. The runtime will prefer the VTID if present
  @VTID(240)
  boolean getConnectionsDisabled();


  /**
   */

  @DISPID(2537) //= 0x9e9. The runtime will prefer the VTID if present
  @VTID(241)
  void enableConnections();


  /**
   * <p>
   * Getter method for the COM property "ShowPivotChartActiveFields"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2538) //= 0x9ea. The runtime will prefer the VTID if present
  @VTID(242)
  boolean getShowPivotChartActiveFields();


  /**
   * <p>
   * Setter method for the COM property "ShowPivotChartActiveFields"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2538) //= 0x9ea. The runtime will prefer the VTID if present
  @VTID(243)
  void setShowPivotChartActiveFields(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(244)
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr);


  /**
   * <p>
   * Getter method for the COM property "IconSets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.IconSets
   */

  @DISPID(2539) //= 0x9eb. The runtime will prefer the VTID if present
  @VTID(245)
  com.exceljava.com4j.excel.IconSets getIconSets();


  /**
   * <p>
   * Getter method for the COM property "EncryptionProvider"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2540) //= 0x9ec. The runtime will prefer the VTID if present
  @VTID(246)
  java.lang.String getEncryptionProvider();


  /**
   * <p>
   * Setter method for the COM property "EncryptionProvider"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2540) //= 0x9ec. The runtime will prefer the VTID if present
  @VTID(247)
  void setEncryptionProvider(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DoNotPromptForConvert"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2541) //= 0x9ed. The runtime will prefer the VTID if present
  @VTID(248)
  boolean getDoNotPromptForConvert();


  /**
   * <p>
   * Setter method for the COM property "DoNotPromptForConvert"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2541) //= 0x9ed. The runtime will prefer the VTID if present
  @VTID(249)
  void setDoNotPromptForConvert(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ForceFullCalculation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2542) //= 0x9ee. The runtime will prefer the VTID if present
  @VTID(250)
  boolean getForceFullCalculation();


  /**
   * <p>
   * Setter method for the COM property "ForceFullCalculation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2542) //= 0x9ee. The runtime will prefer the VTID if present
  @VTID(251)
  void setForceFullCalculation(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protectSharing();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sharingPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, password, writeResPassword, readOnlyRecommended, createBackup, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protectSharing(filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sharingPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sharingPassword);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sharingPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2543) //= 0x9ef. The runtime will prefer the VTID if present
  @VTID(252)
  void protectSharing(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sharingPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);


  /**
   * <p>
   * Getter method for the COM property "SlicerCaches"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCaches
   */

  @DISPID(2866) //= 0xb32. The runtime will prefer the VTID if present
  @VTID(253)
  com.exceljava.com4j.excel.SlicerCaches getSlicerCaches();


  /**
   * <p>
   * Getter method for the COM property "ActiveSlicer"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Slicer
   */

  @DISPID(2867) //= 0xb33. The runtime will prefer the VTID if present
  @VTID(254)
  com.exceljava.com4j.excel.Slicer getActiveSlicer();


  /**
   * <p>
   * Getter method for the COM property "DefaultSlicerStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2868) //= 0xb34. The runtime will prefer the VTID if present
  @VTID(255)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDefaultSlicerStyle();


  /**
   * <p>
   * Setter method for the COM property "DefaultSlicerStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2868) //= 0xb34. The runtime will prefer the VTID if present
  @VTID(256)
  void setDefaultSlicerStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @DISPID(2869) //= 0xb35. The runtime will prefer the VTID if present
  @VTID(257)
  void dummy26();


  /**
   */

  @DISPID(2870) //= 0xb36. The runtime will prefer the VTID if present
  @VTID(258)
  void dummy27();


  /**
   * <p>
   * Getter method for the COM property "AccuracyVersion"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2871) //= 0xb37. The runtime will prefer the VTID if present
  @VTID(259)
  int getAccuracyVersion();


  /**
   * <p>
   * Setter method for the COM property "AccuracyVersion"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2871) //= 0xb37. The runtime will prefer the VTID if present
  @VTID(260)
  void setAccuracyVersion(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "CaseSensitive"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3056) //= 0xbf0. The runtime will prefer the VTID if present
  @VTID(261)
  boolean getCaseSensitive();


  /**
   * <p>
   * Getter method for the COM property "UseWholeCellCriteria"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3057) //= 0xbf1. The runtime will prefer the VTID if present
  @VTID(262)
  boolean getUseWholeCellCriteria();


  /**
   * <p>
   * Getter method for the COM property "UseWildcards"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3058) //= 0xbf2. The runtime will prefer the VTID if present
  @VTID(263)
  boolean getUseWildcards();


  /**
   * <p>
   * Getter method for the COM property "PivotTables"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(690) //= 0x2b2. The runtime will prefer the VTID if present
  @VTID(264)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getPivotTables();


  /**
   * <p>
   * Getter method for the COM property "Model"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Model
   */

  @DISPID(3059) //= 0xbf3. The runtime will prefer the VTID if present
  @VTID(265)
  com.exceljava.com4j.excel.Model getModel();


  /**
   * <p>
   * Getter method for the COM property "ChartDataPointTrack"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2998) //= 0xbb6. The runtime will prefer the VTID if present
  @VTID(266)
  boolean getChartDataPointTrack();


  /**
   * <p>
   * Setter method for the COM property "ChartDataPointTrack"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2998) //= 0xbb6. The runtime will prefer the VTID if present
  @VTID(267)
  void setChartDataPointTrack(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DefaultTimelineStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3060) //= 0xbf4. The runtime will prefer the VTID if present
  @VTID(268)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDefaultTimelineStyle();


  /**
   * <p>
   * Setter method for the COM property "DefaultTimelineStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3060) //= 0xbf4. The runtime will prefer the VTID if present
  @VTID(269)
  void setDefaultTimelineStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Queries"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Queries
   */

  @DISPID(3186) //= 0xc72. The runtime will prefer the VTID if present
  @VTID(270)
  com.exceljava.com4j.excel.Queries getQueries();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter forecastStart is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter forecastEnd is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter confInt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seasonality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataCompletion is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter forecastEnd is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter confInt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seasonality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataCompletion is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter confInt is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seasonality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataCompletion is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter seasonality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataCompletion is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, confInt, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataCompletion is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, confInt, seasonality, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seasonality Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seasonality);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter aggregation is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, confInt, seasonality, dataCompletion, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seasonality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataCompletion Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seasonality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataCompletion);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, confInt, seasonality, dataCompletion, aggregation, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seasonality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataCompletion Optional parameter. Default value is com4j.Variant.getMissing()
   * @param aggregation Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seasonality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataCompletion,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object aggregation);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showStatsTable is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createForecastSheet(timeline, values, forecastStart, forecastEnd, confInt, seasonality, dataCompletion, aggregation, chartType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seasonality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataCompletion Optional parameter. Default value is com4j.Variant.getMissing()
   * @param aggregation Optional parameter. Default value is com4j.Variant.getMissing()
   * @param chartType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seasonality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataCompletion,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object aggregation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object chartType);

  /**
   * @param timeline Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param values Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param forecastStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @param forecastEnd Optional parameter. Default value is com4j.Variant.getMissing()
   * @param confInt Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seasonality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataCompletion Optional parameter. Default value is com4j.Variant.getMissing()
   * @param aggregation Optional parameter. Default value is com4j.Variant.getMissing()
   * @param chartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showStatsTable Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3167) //= 0xc5f. The runtime will prefer the VTID if present
  @VTID(271)
  void createForecastSheet(
    com.exceljava.com4j.excel.Range timeline,
    com.exceljava.com4j.excel.Range values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastStart,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object forecastEnd,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object confInt,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seasonality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataCompletion,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object aggregation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object chartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showStatsTable);


  /**
   * <p>
   * Getter method for the COM property "WorkIdentity"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3173) //= 0xc65. The runtime will prefer the VTID if present
  @VTID(272)
  java.lang.String getWorkIdentity();


  /**
   * <p>
   * Setter method for the COM property "WorkIdentity"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(3173) //= 0xc65. The runtime will prefer the VTID if present
  @VTID(273)
  void setWorkIdentity(
    java.lang.String rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSaveAsAccessMode parameter accessMode is set to 1</li><li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13}, javaType = {com.exceljava.com4j.excel.XlSaveAsAccessMode.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter conflictResolution is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local, workIdentity, 1033);
   * </code>
   * </p>
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param workIdentity Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object workIdentity);

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param accessMode Optional parameter. Default value is 1
   * @param conflictResolution Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param workIdentity Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(274)
  void saveAs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSaveAsAccessMode accessMode,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object conflictResolution,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object workIdentity,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   * @param workIdentity Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(275)
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object workIdentity);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwriteUrl is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * publishToDocs(title, disclosureScope, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param disclosureScope Mandatory com.exceljava.com4j.excel.XlPublishToDocsDisclosureScope parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3185) //= 0xc71. The runtime will prefer the VTID if present
  @VTID(276)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  java.lang.String publishToDocs(
    java.lang.String title,
    com.exceljava.com4j.excel.XlPublishToDocsDisclosureScope disclosureScope);

  /**
   * @param title Mandatory java.lang.String parameter.
   * @param disclosureScope Mandatory com.exceljava.com4j.excel.XlPublishToDocsDisclosureScope parameter.
   * @param overwriteUrl Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3185) //= 0xc71. The runtime will prefer the VTID if present
  @VTID(276)
  java.lang.String publishToDocs(
    java.lang.String title,
    com.exceljava.com4j.excel.XlPublishToDocsDisclosureScope disclosureScope,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwriteUrl);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * lookUpInDocs(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishedDocs
   */

  @DISPID(3227) //= 0xc9b. The runtime will prefer the VTID if present
  @VTID(277)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.PublishedDocs lookUpInDocs();

  /**
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PublishedDocs
   */

  @DISPID(3227) //= 0xc9b. The runtime will prefer the VTID if present
  @VTID(277)
  com.exceljava.com4j.excel.PublishedDocs lookUpInDocs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter publishType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameConflict is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter bstrGroupName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * publishToPBI(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3257) //= 0xcb9. The runtime will prefer the VTID if present
  @VTID(278)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=3)
  java.lang.String publishToPBI();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter nameConflict is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter bstrGroupName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * publishToPBI(publishType, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param publishType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3257) //= 0xcb9. The runtime will prefer the VTID if present
  @VTID(278)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  java.lang.String publishToPBI(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object publishType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter bstrGroupName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * publishToPBI(publishType, nameConflict, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param publishType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameConflict Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3257) //= 0xcb9. The runtime will prefer the VTID if present
  @VTID(278)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  java.lang.String publishToPBI(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object publishType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameConflict);

  /**
   * @param publishType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameConflict Optional parameter. Default value is com4j.Variant.getMissing()
   * @param bstrGroupName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3257) //= 0xcb9. The runtime will prefer the VTID if present
  @VTID(278)
  java.lang.String publishToPBI(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object publishType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameConflict,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object bstrGroupName);


  /**
   * <p>
   * Getter method for the COM property "AutoSaveOn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3232) //= 0xca0. The runtime will prefer the VTID if present
  @VTID(279)
  boolean getAutoSaveOn();


  /**
   * <p>
   * Setter method for the COM property "AutoSaveOn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3232) //= 0xca0. The runtime will prefer the VTID if present
  @VTID(280)
  void setAutoSaveOn(
    boolean rhs);


  // Properties:
}
