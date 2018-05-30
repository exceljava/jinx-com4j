package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E16E-0000-0000-C000-000000000046}")
public interface _CodeModule extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponent
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.vbide._VBComponent getParent();


  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide.VBE getVBE();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param pbstrName Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  void setName(
    java.lang.String pbstrName);


  /**
   * @param string Mandatory java.lang.String parameter.
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  void addFromString(
    java.lang.String string);


  /**
   * @param fileName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  void addFromFile(
    java.lang.String fileName);


  /**
   * <p>
   * Getter method for the COM property "Lines"
   * </p>
   * @param startLine Mandatory int parameter.
   * @param count Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String getLines(
    int startLine,
    int count);


  /**
   * <p>
   * Getter method for the COM property "CountOfLines"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  int getCountOfLines();


  /**
   * @param line Mandatory int parameter.
   * @param string Mandatory java.lang.String parameter.
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  void insertLines(
    int line,
    java.lang.String string);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter count is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * deleteLines(startLine, 1);
   * </code>
   * </p>
   * @param startLine Mandatory int parameter.
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  void deleteLines(
    int startLine);

  /**
   * @param startLine Mandatory int parameter.
   * @param count Optional parameter. Default value is 1
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  void deleteLines(
    int startLine,
    @Optional @DefaultValue("1") int count);


  /**
   * @param line Mandatory int parameter.
   * @param string Mandatory java.lang.String parameter.
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  void replaceLine(
    int line,
    java.lang.String string);


  /**
   * <p>
   * Getter method for the COM property "ProcStartLine"
   * </p>
   * @param procName Mandatory java.lang.String parameter.
   * @param procKind Mandatory com.exceljava.com4j.vbide.vbext_ProcKind parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  int getProcStartLine(
    java.lang.String procName,
    com.exceljava.com4j.vbide.vbext_ProcKind procKind);


  /**
   * <p>
   * Getter method for the COM property "ProcCountLines"
   * </p>
   * @param procName Mandatory java.lang.String parameter.
   * @param procKind Mandatory com.exceljava.com4j.vbide.vbext_ProcKind parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(19)
  int getProcCountLines(
    java.lang.String procName,
    com.exceljava.com4j.vbide.vbext_ProcKind procKind);


  /**
   * <p>
   * Getter method for the COM property "ProcBodyLine"
   * </p>
   * @param procName Mandatory java.lang.String parameter.
   * @param procKind Mandatory com.exceljava.com4j.vbide.vbext_ProcKind parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743821) //= 0x6002000d. The runtime will prefer the VTID if present
  @VTID(20)
  int getProcBodyLine(
    java.lang.String procName,
    com.exceljava.com4j.vbide.vbext_ProcKind procKind);


  /**
   * <p>
   * Getter method for the COM property "ProcOfLine"
   * </p>
   * @param line Mandatory int parameter.
   * @param procKind Mandatory Holder<com.exceljava.com4j.vbide.vbext_ProcKind> parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getProcOfLine(
    int line,
    Holder<com.exceljava.com4j.vbide.vbext_ProcKind> procKind);


  /**
   * <p>
   * Getter method for the COM property "CountOfDeclarationLines"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743823) //= 0x6002000f. The runtime will prefer the VTID if present
  @VTID(22)
  int getCountOfDeclarationLines();


  /**
   * @param eventName Mandatory java.lang.String parameter.
   * @param objectName Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(23)
  int createEventProc(
    java.lang.String eventName,
    java.lang.String objectName);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter wholeWord is set to false</li><li>boolean parameter matchCase is set to false</li><li>boolean parameter patternSearch is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(target, startLine, startColumn, endLine, endColumn, false, false, false);
   * </code>
   * </p>
   * @param target Mandatory java.lang.String parameter.
   * @param startLine Mandatory Holder<Integer> parameter.
   * @param startColumn Mandatory Holder<Integer> parameter.
   * @param endLine Mandatory Holder<Integer> parameter.
   * @param endColumn Mandatory Holder<Integer> parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {boolean.class, boolean.class, boolean.class}, nativeType = {NativeType.VariantBool, NativeType.VariantBool, NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL, Variant.Type.VT_BOOL, Variant.Type.VT_BOOL}, literal = {"false", "false", "false"})
  @ReturnValue(index=8)
  boolean find(
    java.lang.String target,
    Holder<Integer> startLine,
    Holder<Integer> startColumn,
    Holder<Integer> endLine,
    Holder<Integer> endColumn);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter matchCase is set to false</li><li>boolean parameter patternSearch is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(target, startLine, startColumn, endLine, endColumn, wholeWord, false, false);
   * </code>
   * </p>
   * @param target Mandatory java.lang.String parameter.
   * @param startLine Mandatory Holder<Integer> parameter.
   * @param startColumn Mandatory Holder<Integer> parameter.
   * @param endLine Mandatory Holder<Integer> parameter.
   * @param endColumn Mandatory Holder<Integer> parameter.
   * @param wholeWord Optional parameter. Default value is false
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {boolean.class, boolean.class}, nativeType = {NativeType.VariantBool, NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL, Variant.Type.VT_BOOL}, literal = {"false", "false"})
  @ReturnValue(index=8)
  boolean find(
    java.lang.String target,
    Holder<Integer> startLine,
    Holder<Integer> startColumn,
    Holder<Integer> endLine,
    Holder<Integer> endColumn,
    @Optional @DefaultValue("0") boolean wholeWord);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter patternSearch is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, false);
   * </code>
   * </p>
   * @param target Mandatory java.lang.String parameter.
   * @param startLine Mandatory Holder<Integer> parameter.
   * @param startColumn Mandatory Holder<Integer> parameter.
   * @param endLine Mandatory Holder<Integer> parameter.
   * @param endColumn Mandatory Holder<Integer> parameter.
   * @param wholeWord Optional parameter. Default value is false
   * @param matchCase Optional parameter. Default value is false
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  @ReturnValue(index=8)
  boolean find(
    java.lang.String target,
    Holder<Integer> startLine,
    Holder<Integer> startColumn,
    Holder<Integer> endLine,
    Holder<Integer> endColumn,
    @Optional @DefaultValue("0") boolean wholeWord,
    @Optional @DefaultValue("0") boolean matchCase);

  /**
   * @param target Mandatory java.lang.String parameter.
   * @param startLine Mandatory Holder<Integer> parameter.
   * @param startColumn Mandatory Holder<Integer> parameter.
   * @param endLine Mandatory Holder<Integer> parameter.
   * @param endColumn Mandatory Holder<Integer> parameter.
   * @param wholeWord Optional parameter. Default value is false
   * @param matchCase Optional parameter. Default value is false
   * @param patternSearch Optional parameter. Default value is false
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(24)
  boolean find(
    java.lang.String target,
    Holder<Integer> startLine,
    Holder<Integer> startColumn,
    Holder<Integer> endLine,
    Holder<Integer> endColumn,
    @Optional @DefaultValue("0") boolean wholeWord,
    @Optional @DefaultValue("0") boolean matchCase,
    @Optional @DefaultValue("0") boolean patternSearch);


  /**
   * <p>
   * Getter method for the COM property "CodePane"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._CodePane
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.vbide._CodePane getCodePane();


  // Properties:
}
