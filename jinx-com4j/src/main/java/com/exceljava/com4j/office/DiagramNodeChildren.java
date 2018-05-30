package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C036F-0000-0000-C000-000000000046}")
public interface DiagramNodeChildren extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(9)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.office.DiagramNode item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to -1</li><li>com.exceljava.com4j.office.MsoDiagramNodeType parameter nodeType is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(-1, 1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, com.exceljava.com4j.office.MsoDiagramNodeType.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I2, Variant.Type.VT_I4}, literal = {"-1", "1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.DiagramNode addNode();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoDiagramNodeType parameter nodeType is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(index, 1);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.MsoDiagramNodeType.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.DiagramNode addNode(
    @Optional @DefaultValue("-1") @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is -1
   * @param nodeType Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.DiagramNode addNode(
    @Optional @DefaultValue("-1") @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoDiagramNodeType nodeType);


  /**
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(12)
  void selectAll();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(14)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "FirstChild"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.DiagramNode getFirstChild();


  /**
   * <p>
   * Getter method for the COM property "LastChild"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.DiagramNode getLastChild();


  // Properties:
}
