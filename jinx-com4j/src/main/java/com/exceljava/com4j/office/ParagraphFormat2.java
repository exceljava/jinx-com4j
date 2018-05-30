package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0399-0000-0000-C000-000000000046}")
public interface ParagraphFormat2 extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "Alignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoParagraphAlignment
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.MsoParagraphAlignment getAlignment();


  /**
   * <p>
   * Setter method for the COM property "Alignment"
   * </p>
   * @param alignment Mandatory com.exceljava.com4j.office.MsoParagraphAlignment parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void setAlignment(
    com.exceljava.com4j.office.MsoParagraphAlignment alignment);


  /**
   * <p>
   * Getter method for the COM property "BaselineAlignment"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBaselineAlignment
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.MsoBaselineAlignment getBaselineAlignment();


  /**
   * <p>
   * Setter method for the COM property "BaselineAlignment"
   * </p>
   * @param baselineAlignment Mandatory com.exceljava.com4j.office.MsoBaselineAlignment parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  void setBaselineAlignment(
    com.exceljava.com4j.office.MsoBaselineAlignment baselineAlignment);


  /**
   * <p>
   * Getter method for the COM property "Bullet"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.BulletFormat2
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.BulletFormat2 getBullet();


  /**
   * <p>
   * Getter method for the COM property "FarEastLineBreakLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.MsoTriState getFarEastLineBreakLevel();


  /**
   * <p>
   * Setter method for the COM property "FarEastLineBreakLevel"
   * </p>
   * @param _break Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  void setFarEastLineBreakLevel(
    com.exceljava.com4j.office.MsoTriState _break);


  /**
   * <p>
   * Getter method for the COM property "FirstLineIndent"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  float getFirstLineIndent();


  /**
   * <p>
   * Setter method for the COM property "FirstLineIndent"
   * </p>
   * @param indent Mandatory float parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(18)
  void setFirstLineIndent(
    float indent);


  /**
   * <p>
   * Getter method for the COM property "HangingPunctuation"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoTriState getHangingPunctuation();


  /**
   * <p>
   * Setter method for the COM property "HangingPunctuation"
   * </p>
   * @param hanging Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(20)
  void setHangingPunctuation(
    com.exceljava.com4j.office.MsoTriState hanging);


  /**
   * <p>
   * Getter method for the COM property "IndentLevel"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(21)
  int getIndentLevel();


  /**
   * <p>
   * Setter method for the COM property "IndentLevel"
   * </p>
   * @param level Mandatory int parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(22)
  void setIndentLevel(
    int level);


  /**
   * <p>
   * Getter method for the COM property "LeftIndent"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(23)
  float getLeftIndent();


  /**
   * <p>
   * Setter method for the COM property "LeftIndent"
   * </p>
   * @param indent Mandatory float parameter.
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(24)
  void setLeftIndent(
    float indent);


  /**
   * <p>
   * Getter method for the COM property "LineRuleAfter"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.MsoTriState getLineRuleAfter();


  /**
   * <p>
   * Setter method for the COM property "LineRuleAfter"
   * </p>
   * @param lineRule Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(26)
  void setLineRuleAfter(
    com.exceljava.com4j.office.MsoTriState lineRule);


  /**
   * <p>
   * Getter method for the COM property "LineRuleBefore"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.MsoTriState getLineRuleBefore();


  /**
   * <p>
   * Setter method for the COM property "LineRuleBefore"
   * </p>
   * @param lineRule Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(28)
  void setLineRuleBefore(
    com.exceljava.com4j.office.MsoTriState lineRule);


  /**
   * <p>
   * Getter method for the COM property "LineRuleWithin"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(29)
  com.exceljava.com4j.office.MsoTriState getLineRuleWithin();


  /**
   * <p>
   * Setter method for the COM property "LineRuleWithin"
   * </p>
   * @param lineRule Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(30)
  void setLineRuleWithin(
    com.exceljava.com4j.office.MsoTriState lineRule);


  /**
   * <p>
   * Getter method for the COM property "RightIndent"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(31)
  float getRightIndent();


  /**
   * <p>
   * Setter method for the COM property "RightIndent"
   * </p>
   * @param indent Mandatory float parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(32)
  void setRightIndent(
    float indent);


  /**
   * <p>
   * Getter method for the COM property "SpaceAfter"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(33)
  float getSpaceAfter();


  /**
   * <p>
   * Setter method for the COM property "SpaceAfter"
   * </p>
   * @param space Mandatory float parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(34)
  void setSpaceAfter(
    float space);


  /**
   * <p>
   * Getter method for the COM property "SpaceBefore"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(35)
  float getSpaceBefore();


  /**
   * <p>
   * Setter method for the COM property "SpaceBefore"
   * </p>
   * @param space Mandatory float parameter.
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(36)
  void setSpaceBefore(
    float space);


  /**
   * <p>
   * Getter method for the COM property "SpaceWithin"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(37)
  float getSpaceWithin();


  /**
   * <p>
   * Setter method for the COM property "SpaceWithin"
   * </p>
   * @param space Mandatory float parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(38)
  void setSpaceWithin(
    float space);


  /**
   * <p>
   * Getter method for the COM property "TabStops"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TabStops2
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(39)
  com.exceljava.com4j.office.TabStops2 getTabStops();


  @VTID(39)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.TabStops2.class})
  com.exceljava.com4j.office.TabStop2 getTabStops(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "TextDirection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTextDirection
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(40)
  com.exceljava.com4j.office.MsoTextDirection getTextDirection();


  /**
   * <p>
   * Setter method for the COM property "TextDirection"
   * </p>
   * @param direction Mandatory com.exceljava.com4j.office.MsoTextDirection parameter.
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(41)
  void setTextDirection(
    com.exceljava.com4j.office.MsoTextDirection direction);


  /**
   * <p>
   * Getter method for the COM property "WordWrap"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(42)
  com.exceljava.com4j.office.MsoTriState getWordWrap();


  /**
   * <p>
   * Setter method for the COM property "WordWrap"
   * </p>
   * @param wordWrap Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(43)
  void setWordWrap(
    com.exceljava.com4j.office.MsoTriState wordWrap);


  // Properties:
}
