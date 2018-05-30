package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0340-0000-0000-C000-000000000046}")
public interface Scripts extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.office.Script item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com4j.Com4jObject parameter anchor is set to null</li><li>com.exceljava.com4j.office.MsoScriptLocation parameter location is set to 2</li><li>com.exceljava.com4j.office.MsoScriptLanguage parameter language is set to 2</li><li>java.lang.String parameter id is set to ""</li><li>java.lang.String parameter extended is set to ""</li><li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(null, 2, 2, "", "", "");
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {com4j.Com4jObject.class, com.exceljava.com4j.office.MsoScriptLocation.class, com.exceljava.com4j.office.MsoScriptLanguage.class, java.lang.String.class, java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.Dispatch, NativeType.Int32, NativeType.Int32, NativeType.BSTR, NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_DISPATCH, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"null", "2", "2", "", "", ""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoScriptLocation parameter location is set to 2</li><li>com.exceljava.com4j.office.MsoScriptLanguage parameter language is set to 2</li><li>java.lang.String parameter id is set to ""</li><li>java.lang.String parameter extended is set to ""</li><li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, 2, 2, "", "", "");
   * </code>
   * </p>
   * @param anchor Optional parameter. Default value is unprintable.
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.office.MsoScriptLocation.class, com.exceljava.com4j.office.MsoScriptLanguage.class, java.lang.String.class, java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.BSTR, NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"2", "2", "", "", ""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoScriptLanguage parameter language is set to 2</li><li>java.lang.String parameter id is set to ""</li><li>java.lang.String parameter extended is set to ""</li><li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, location, 2, "", "", "");
   * </code>
   * </p>
   * @param anchor Optional parameter. Default value is unprintable.
   * @param location Optional parameter. Default value is 2
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {com.exceljava.com4j.office.MsoScriptLanguage.class, java.lang.String.class, java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.Int32, NativeType.BSTR, NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"2", "", "", ""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLocation location);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter id is set to ""</li><li>java.lang.String parameter extended is set to ""</li><li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, location, language, "", "", "");
   * </code>
   * </p>
   * @param anchor Optional parameter. Default value is unprintable.
   * @param location Optional parameter. Default value is 2
   * @param language Optional parameter. Default value is 2
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.String.class, java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"", "", ""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLocation location,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLanguage language);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter extended is set to ""</li><li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, location, language, id, "", "");
   * </code>
   * </p>
   * @param anchor Optional parameter. Default value is unprintable.
   * @param location Optional parameter. Default value is 2
   * @param language Optional parameter. Default value is 2
   * @param id Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"", ""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLocation location,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLanguage language,
    @Optional @DefaultValue("") java.lang.String id);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter scriptText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, location, language, id, extended, "");
   * </code>
   * </p>
   * @param anchor Optional parameter. Default value is unprintable.
   * @param location Optional parameter. Default value is 2
   * @param language Optional parameter. Default value is 2
   * @param id Optional parameter. Default value is ""
   * @param extended Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLocation location,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLanguage language,
    @Optional @DefaultValue("") java.lang.String id,
    @Optional @DefaultValue("") java.lang.String extended);

  /**
   * @param anchor Optional parameter. Default value is unprintable.
   * @param location Optional parameter. Default value is 2
   * @param language Optional parameter. Default value is 2
   * @param id Optional parameter. Default value is ""
   * @param extended Optional parameter. Default value is ""
   * @param scriptText Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.Script add(
    @Optional @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLocation location,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoScriptLanguage language,
    @Optional @DefaultValue("") java.lang.String id,
    @Optional @DefaultValue("") java.lang.String extended,
    @Optional @DefaultValue("") java.lang.String scriptText);


  /**
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  void delete();


  // Properties:
}
