package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRemoveDocInfoType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRDIComments(1),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlRDIRemovePersonalInformation(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlRDIEmailHeader(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlRDIRoutingSlip(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlRDISendForReview(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlRDIDocumentProperties(8),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlRDIDocumentWorkspace(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlRDIInkAnnotations(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlRDIScenarioComments(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlRDIPublishInfo(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlRDIDocumentServerProperties(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlRDIDocumentManagementPolicy(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlRDIContentType(16),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlRDIDefinedNameComments(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlRDIInactiveDataConnections(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlRDIPrinterPath(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlRDIInlineWebExtensions(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlRDITaskpaneWebExtensions(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlRDIExcelDataModel(23),
  /**
   * <p>
   * The value of this constant is 99
   * </p>
   */
  xlRDIAll(99),
  ;

  private final int value;
  XlRemoveDocInfoType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
