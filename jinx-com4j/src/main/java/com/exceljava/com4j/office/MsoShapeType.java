package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoShapeType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoShapeTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAutoShape(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCallout(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoChart(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoComment(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoFreeform(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoGroup(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoEmbeddedOLEObject(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoFormControl(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoLine(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoLinkedOLEObject(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoLinkedPicture(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoOLEControlObject(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoPicture(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoPlaceholder(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoTextEffect(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoMedia(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoTextBox(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoScriptAnchor(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoTable(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoCanvas(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoDiagram(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoInk(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoInkComment(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoSmartArt(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoSlicer(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoWebVideo(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  msoContentApp(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  msoGraphic(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  msoLinkedGraphic(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  mso3DModel(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  msoLinked3DModel(31),
  ;

  private final int value;
  MsoShapeType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
