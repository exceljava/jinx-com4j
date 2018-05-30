package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E176-0000-0000-C000-000000000046}")
public interface _CodePane extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Collection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._CodePanes
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.vbide._CodePanes getCollection();


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
   * Getter method for the COM property "Window"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Window
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  com.exceljava.com4j.vbide.Window getWindow();


  /**
   * @param startLine Mandatory Holder<Integer> parameter.
   * @param startColumn Mandatory Holder<Integer> parameter.
   * @param endLine Mandatory Holder<Integer> parameter.
   * @param endColumn Mandatory Holder<Integer> parameter.
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  void getSelection(
    Holder<Integer> startLine,
    Holder<Integer> startColumn,
    Holder<Integer> endLine,
    Holder<Integer> endColumn);


  /**
   * @param startLine Mandatory int parameter.
   * @param startColumn Mandatory int parameter.
   * @param endLine Mandatory int parameter.
   * @param endColumn Mandatory int parameter.
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  void setSelection(
    int startLine,
    int startColumn,
    int endLine,
    int endColumn);


  /**
   * <p>
   * Getter method for the COM property "TopLine"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  int getTopLine();


  /**
   * <p>
   * Setter method for the COM property "TopLine"
   * </p>
   * @param topLine Mandatory int parameter.
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(13)
  void setTopLine(
    int topLine);


  /**
   * <p>
   * Getter method for the COM property "CountOfVisibleLines"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  int getCountOfVisibleLines();


  /**
   * <p>
   * Getter method for the COM property "CodeModule"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._CodeModule
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide._CodeModule getCodeModule();


  /**
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  void show();


  /**
   * <p>
   * Getter method for the COM property "CodePaneView"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_CodePaneview
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.vbide.vbext_CodePaneview getCodePaneView();


  // Properties:
}
