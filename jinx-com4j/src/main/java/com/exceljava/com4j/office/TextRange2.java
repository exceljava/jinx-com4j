package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0397-0000-0000-C000-000000000046}")
public interface TextRange2 extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param pbstrText Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  void setText(
    java.lang.String pbstrText);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(11)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.TextRange2 item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Paragraphs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getParagraphs(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getParagraphs();

  /**
   * <p>
   * Getter method for the COM property "Paragraphs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getParagraphs(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getParagraphs(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Paragraphs"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.TextRange2 getParagraphs(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "Sentences"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSentences(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getSentences();

  /**
   * <p>
   * Getter method for the COM property "Sentences"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSentences(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getSentences(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Sentences"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.TextRange2 getSentences(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "Words"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWords(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getWords();

  /**
   * <p>
   * Getter method for the COM property "Words"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWords(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getWords(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Words"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.TextRange2 getWords(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getCharacters();

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getCharacters(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.TextRange2 getCharacters(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "Lines"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getLines(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getLines();

  /**
   * <p>
   * Getter method for the COM property "Lines"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getLines(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getLines(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Lines"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.TextRange2 getLines(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "Runs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRuns(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getRuns();

  /**
   * <p>
   * Getter method for the COM property "Runs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRuns(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getRuns(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "Runs"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.TextRange2 getRuns(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * Getter method for the COM property "ParagraphFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ParagraphFormat2
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.ParagraphFormat2 getParagraphFormat();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Font2
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.Font2 getFont();


  /**
   * <p>
   * Getter method for the COM property "Length"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(23)
  int getLength();


  /**
   * <p>
   * Getter method for the COM property "Start"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(24)
  int getStart();


  /**
   * <p>
   * Getter method for the COM property "BoundLeft"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(25)
  float getBoundLeft();


  /**
   * <p>
   * Getter method for the COM property "BoundTop"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(26)
  float getBoundTop();


  /**
   * <p>
   * Getter method for the COM property "BoundWidth"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(27)
  float getBoundWidth();


  /**
   * <p>
   * Getter method for the COM property "BoundHeight"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(28)
  float getBoundHeight();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(29)
  com.exceljava.com4j.office.TextRange2 trimText();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter newText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertAfter("");
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.TextRange2 insertAfter();

  /**
   * @param newText Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.TextRange2 insertAfter(
    @Optional @DefaultValue("") java.lang.String newText);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter newText is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertBefore("");
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.TextRange2 insertBefore();

  /**
   * @param newText Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(31)
  com.exceljava.com4j.office.TextRange2 insertBefore(
    @Optional @DefaultValue("") java.lang.String newText);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoTriState parameter unicode is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertSymbol(fontName, charNumber, 0);
   * </code>
   * </p>
   * @param fontName Mandatory java.lang.String parameter.
   * @param charNumber Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.TextRange2 insertSymbol(
    java.lang.String fontName,
    int charNumber);

  /**
   * @param fontName Mandatory java.lang.String parameter.
   * @param charNumber Mandatory int parameter.
   * @param unicode Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(32)
  com.exceljava.com4j.office.TextRange2 insertSymbol(
    java.lang.String fontName,
    int charNumber,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState unicode);


  /**
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(33)
  void select();


  /**
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(34)
  void cut();


  /**
   */

  @DISPID(24) //= 0x18. The runtime will prefer the VTID if present
  @VTID(35)
  void copy();


  /**
   */

  @DISPID(25) //= 0x19. The runtime will prefer the VTID if present
  @VTID(36)
  void delete();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(37)
  com.exceljava.com4j.office.TextRange2 paste();


  /**
   * @param format Mandatory com.exceljava.com4j.office.MsoClipboardFormat parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(27) //= 0x1b. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.TextRange2 pasteSpecial(
    com.exceljava.com4j.office.MsoClipboardFormat format);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoTextChangeCase parameter.
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(39)
  void changeCase(
    com.exceljava.com4j.office.MsoTextChangeCase type);


  /**
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(40)
  void addPeriods();


  /**
   */

  @DISPID(30) //= 0x1e. The runtime will prefer the VTID if present
  @VTID(41)
  void removePeriods();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter after is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter matchCase is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(findWhat, 0, 0, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {int.class, com.exceljava.com4j.office.MsoTriState.class, com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.TextRange2 find(
    java.lang.String findWhat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoTriState parameter matchCase is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(findWhat, after, 0, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.office.MsoTriState.class, com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.TextRange2 find(
    java.lang.String findWhat,
    @Optional @DefaultValue("0") int after);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * find(findWhat, after, matchCase, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @param matchCase Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.TextRange2 find(
    java.lang.String findWhat,
    @Optional @DefaultValue("0") int after,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState matchCase);

  /**
   * @param findWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @param matchCase Optional parameter. Default value is 0
   * @param wholeWords Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(42)
  com.exceljava.com4j.office.TextRange2 find(
    java.lang.String findWhat,
    @Optional @DefaultValue("0") int after,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState matchCase,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState wholeWords);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter after is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter matchCase is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(findWhat, replaceWhat, 0, 0, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @param replaceWhat Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(32) //= 0x20. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {int.class, com.exceljava.com4j.office.MsoTriState.class, com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0", "0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.TextRange2 replace(
    java.lang.String findWhat,
    java.lang.String replaceWhat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoTriState parameter matchCase is set to 0</li><li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(findWhat, replaceWhat, after, 0, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @param replaceWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(32) //= 0x20. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {com.exceljava.com4j.office.MsoTriState.class, com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.TextRange2 replace(
    java.lang.String findWhat,
    java.lang.String replaceWhat,
    @Optional @DefaultValue("0") int after);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoTriState parameter wholeWords is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replace(findWhat, replaceWhat, after, matchCase, 0);
   * </code>
   * </p>
   * @param findWhat Mandatory java.lang.String parameter.
   * @param replaceWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @param matchCase Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(32) //= 0x20. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {com.exceljava.com4j.office.MsoTriState.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.TextRange2 replace(
    java.lang.String findWhat,
    java.lang.String replaceWhat,
    @Optional @DefaultValue("0") int after,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState matchCase);

  /**
   * @param findWhat Mandatory java.lang.String parameter.
   * @param replaceWhat Mandatory java.lang.String parameter.
   * @param after Optional parameter. Default value is 0
   * @param matchCase Optional parameter. Default value is 0
   * @param wholeWords Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(32) //= 0x20. The runtime will prefer the VTID if present
  @VTID(43)
  com.exceljava.com4j.office.TextRange2 replace(
    java.lang.String findWhat,
    java.lang.String replaceWhat,
    @Optional @DefaultValue("0") int after,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState matchCase,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoTriState wholeWords);


  /**
   * @param x1 Mandatory Holder<Float> parameter.
   * @param y1 Mandatory Holder<Float> parameter.
   * @param x2 Mandatory Holder<Float> parameter.
   * @param y2 Mandatory Holder<Float> parameter.
   * @param x3 Mandatory Holder<Float> parameter.
   * @param y3 Mandatory Holder<Float> parameter.
   * @param x4 Mandatory Holder<Float> parameter.
   * @param y4 Mandatory Holder<Float> parameter.
   */

  @DISPID(33) //= 0x21. The runtime will prefer the VTID if present
  @VTID(44)
  void rotatedBounds(
    Holder<Float> x1,
    Holder<Float> y1,
    Holder<Float> x2,
    Holder<Float> y2,
    Holder<Float> x3,
    Holder<Float> y3,
    Holder<Float> x4,
    Holder<Float> y4);


  /**
   * <p>
   * Getter method for the COM property "LanguageID"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoLanguageID
   */

  @DISPID(34) //= 0x22. The runtime will prefer the VTID if present
  @VTID(45)
  com.exceljava.com4j.office.MsoLanguageID getLanguageID();


  /**
   * <p>
   * Setter method for the COM property "LanguageID"
   * </p>
   * @param languageID Mandatory com.exceljava.com4j.office.MsoLanguageID parameter.
   */

  @DISPID(34) //= 0x22. The runtime will prefer the VTID if present
  @VTID(46)
  void setLanguageID(
    com.exceljava.com4j.office.MsoLanguageID languageID);


  /**
   */

  @DISPID(35) //= 0x23. The runtime will prefer the VTID if present
  @VTID(47)
  void rtlRun();


  /**
   */

  @DISPID(36) //= 0x24. The runtime will prefer the VTID if present
  @VTID(48)
  void ltrRun();


  /**
   * <p>
   * Getter method for the COM property "MathZones"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter start is set to -1</li><li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getMathZones(-1, -1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(37) //= 0x25. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-1", "-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getMathZones();

  /**
   * <p>
   * Getter method for the COM property "MathZones"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter length is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getMathZones(start, -1);
   * </code>
   * </p>
   * @param start Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(37) //= 0x25. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.TextRange2 getMathZones(
    @Optional @DefaultValue("-1") int start);

  /**
   * <p>
   * Getter method for the COM property "MathZones"
   * </p>
   * @param start Optional parameter. Default value is -1
   * @param length Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(37) //= 0x25. The runtime will prefer the VTID if present
  @VTID(49)
  com.exceljava.com4j.office.TextRange2 getMathZones(
    @Optional @DefaultValue("-1") int start,
    @Optional @DefaultValue("-1") int length);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter formula is set to ""</li><li>int parameter position is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertChartField(chartFieldType, "", -1);
   * </code>
   * </p>
   * @param chartFieldType Mandatory com.exceljava.com4j.office.MsoChartFieldType parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(38) //= 0x26. The runtime will prefer the VTID if present
  @VTID(50)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.String.class, int.class}, nativeType = {NativeType.BSTR, NativeType.Int32}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I4}, literal = {"", "-1"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.TextRange2 insertChartField(
    com.exceljava.com4j.office.MsoChartFieldType chartFieldType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter position is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertChartField(chartFieldType, formula, -1);
   * </code>
   * </p>
   * @param chartFieldType Mandatory com.exceljava.com4j.office.MsoChartFieldType parameter.
   * @param formula Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(38) //= 0x26. The runtime will prefer the VTID if present
  @VTID(50)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.TextRange2 insertChartField(
    com.exceljava.com4j.office.MsoChartFieldType chartFieldType,
    @Optional @DefaultValue("") java.lang.String formula);

  /**
   * @param chartFieldType Mandatory com.exceljava.com4j.office.MsoChartFieldType parameter.
   * @param formula Optional parameter. Default value is ""
   * @param position Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.TextRange2
   */

  @DISPID(38) //= 0x26. The runtime will prefer the VTID if present
  @VTID(50)
  com.exceljava.com4j.office.TextRange2 insertChartField(
    com.exceljava.com4j.office.MsoChartFieldType chartFieldType,
    @Optional @DefaultValue("") java.lang.String formula,
    @Optional @DefaultValue("-1") int position);


  // Properties:
}
