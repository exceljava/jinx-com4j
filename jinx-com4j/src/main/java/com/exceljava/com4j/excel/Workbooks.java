package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208DB-0000-0000-C000-000000000046}")
public interface Workbooks extends Com4jObject,Iterable<Com4jObject> {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter template is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(181) //= 0xb5. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Workbook add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(template, 1033);
   * </code>
   * </p>
   * @param template Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(181) //= 0xb5. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Workbook add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object template);

  /**
   * @param template Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(181) //= 0xb5. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.excel._Workbook add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object template,
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
   * close(1033);
   * </code>
   * </p>
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void close();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(277) //= 0x115. The runtime will prefer the VTID if present
  @VTID(11)
  void close(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(12)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(170) //= 0xaa. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.excel._Workbook getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter updateLinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 14}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 14}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 14}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 14}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 14}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 14}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 14}, optParamIndex = {7, 8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 14}, optParamIndex = {8, 9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 14}, optParamIndex = {9, 10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 14}, optParamIndex = {10, 11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 14}, optParamIndex = {11, 12, 13}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 14}, optParamIndex = {12, 13}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14}, optParamIndex = {13}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=14)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(682) //= 0x2aa. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.excel._Workbook _Open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(683) //= 0x2ab. The runtime will prefer the VTID if present
  @VTID(16)
  void __OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(17)
  @DefaultMethod
  com.exceljava.com4j.excel._Workbook get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15, 16}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, optParamIndex = {16}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1773) //= 0x6ed. The runtime will prefer the VTID if present
  @VTID(18)
  void _OpenText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter updateLinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 16}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 16}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 16}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 16}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 16}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreReadOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 16}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 16}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter delimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 16}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter editable is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 16}, optParamIndex = {9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter notify is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 16}, optParamIndex = {10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter converter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 16}, optParamIndex = {11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 16}, optParamIndex = {12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 16}, optParamIndex = {13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter corruptLoad is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16}, optParamIndex = {14, 15}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
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
   * open(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local, corruptLoad, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param corruptLoad Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16}, optParamIndex = {15}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=16)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object corruptLoad);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreReadOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param delimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param editable Optional parameter. Default value is com4j.Variant.getMissing()
   * @param notify Optional parameter. Default value is com4j.Variant.getMissing()
   * @param converter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param corruptLoad Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(1923) //= 0x783. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.excel._Workbook open(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreReadOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object delimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object editable,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object notify,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object converter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object corruptLoad,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter origin is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter startRow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataType is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, com4j.Variant.getMissing(), 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlTextQualifier parameter textQualifier is set to 1</li><li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {com.exceljava.com4j.excel.XlTextQualifier.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter consecutiveDelimiter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tab is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter semicolon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comma is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter space is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter other is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter otherChar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fieldInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter thousandsSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15, 16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter trailingMinusNumbers is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, optParamIndex = {16, 17, 18}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trailingMinusNumbers Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, optParamIndex = {17, 18}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object trailingMinusNumbers);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openText(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers, local, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trailingMinusNumbers Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}, optParamIndex = {18}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object trailingMinusNumbers,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param origin Optional parameter. Default value is com4j.Variant.getMissing()
   * @param startRow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textQualifier Optional parameter. Default value is 1
   * @param consecutiveDelimiter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tab Optional parameter. Default value is com4j.Variant.getMissing()
   * @param semicolon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comma Optional parameter. Default value is com4j.Variant.getMissing()
   * @param space Optional parameter. Default value is com4j.Variant.getMissing()
   * @param other Optional parameter. Default value is com4j.Variant.getMissing()
   * @param otherChar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fieldInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param thousandsSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param trailingMinusNumbers Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1924) //= 0x784. The runtime will prefer the VTID if present
  @VTID(20)
  void openText(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object origin,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object startRow,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataType,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlTextQualifier textQualifier,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object consecutiveDelimiter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tab,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object semicolon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comma,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object space,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object other,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object otherChar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fieldInfo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object thousandsSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object trailingMinusNumbers,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter commandText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter commandType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importDataAs is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openDatabase(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2067) //= 0x813. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel._Workbook openDatabase(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter commandType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importDataAs is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openDatabase(filename, commandText, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param commandText Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2067) //= 0x813. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel._Workbook openDatabase(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter importDataAs is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openDatabase(filename, commandText, commandType, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param commandText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param commandType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2067) //= 0x813. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel._Workbook openDatabase(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter importDataAs is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openDatabase(filename, commandText, commandType, backgroundQuery, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param commandText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param commandType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2067) //= 0x813. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel._Workbook openDatabase(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param commandText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param commandType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param importDataAs Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2067) //= 0x813. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.excel._Workbook openDatabase(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object commandType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object importDataAs);


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(2069) //= 0x815. The runtime will prefer the VTID if present
  @VTID(22)
  void checkOut(
    java.lang.String filename);


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(2070) //= 0x816. The runtime will prefer the VTID if present
  @VTID(23)
  boolean canCheckOut(
    java.lang.String filename);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter stylesheets is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _OpenXML(filename, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2071) //= 0x817. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Workbook _OpenXML(
    java.lang.String filename);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param stylesheets Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2071) //= 0x817. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.excel._Workbook _OpenXML(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object stylesheets);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter stylesheets is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter loadOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openXML(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2280) //= 0x8e8. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel._Workbook openXML(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter loadOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * openXML(filename, stylesheets, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param stylesheets Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2280) //= 0x8e8. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel._Workbook openXML(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object stylesheets);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param stylesheets Optional parameter. Default value is com4j.Variant.getMissing()
   * @param loadOption Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @DISPID(2280) //= 0x8e8. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.excel._Workbook openXML(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object stylesheets,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object loadOption);


  // Properties:
}
