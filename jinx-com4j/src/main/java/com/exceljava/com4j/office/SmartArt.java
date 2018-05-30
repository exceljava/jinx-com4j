package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C6-0000-0000-C000-000000000046}")
public interface SmartArt extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "AllNodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNodes
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.SmartArtNodes getAllNodes();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SmartArtNodes.class})
  com.exceljava.com4j.office.SmartArtNode getAllNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNodes
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.SmartArtNodes getNodes();


  @VTID(11)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SmartArtNodes.class})
  com.exceljava.com4j.office.SmartArtNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Layout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtLayout
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.SmartArtLayout getLayout();


  /**
   * <p>
   * Setter method for the COM property "Layout"
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  void setLayout(
    com.exceljava.com4j.office.SmartArtLayout layout);


  /**
   * <p>
   * Getter method for the COM property "QuickStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtQuickStyle
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.SmartArtQuickStyle getQuickStyle();


  /**
   * <p>
   * Setter method for the COM property "QuickStyle"
   * </p>
   * @param style Mandatory com.exceljava.com4j.office.SmartArtQuickStyle parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  void setQuickStyle(
    com.exceljava.com4j.office.SmartArtQuickStyle style);


  /**
   * <p>
   * Getter method for the COM property "Color"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtColor
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.SmartArtColor getColor();


  /**
   * <p>
   * Setter method for the COM property "Color"
   * </p>
   * @param colorStyle Mandatory com.exceljava.com4j.office.SmartArtColor parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  void setColor(
    com.exceljava.com4j.office.SmartArtColor colorStyle);


  /**
   * <p>
   * Getter method for the COM property "Reverse"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoTriState getReverse();


  /**
   * <p>
   * Setter method for the COM property "Reverse"
   * </p>
   * @param reverse Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(19)
  void setReverse(
    com.exceljava.com4j.office.MsoTriState reverse);


  /**
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(20)
  void reset();


  // Properties:
}
