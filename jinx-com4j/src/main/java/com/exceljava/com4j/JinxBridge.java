package com.exceljava.com4j;

import com.exceljava.jinx.ExcelAddIn;
import com.exceljava.com4j.excel._Application;
import com4j.COM4J;
import com4j.Com4jObject;
import com4j.ComThread;
import com4j.ROT;

/**
 * Bridge between the Jinx add-in and com4j.
 * Used for obtaining com4j COM wrappers from code running in Excel using Jinx.
 */
public class JinxBridge {

    private static final ThreadLocal<_Application> xlApp = new ThreadLocal<_Application>() {
        public _Application initialValue() {
            return null;
        }
    };

    /**
     * Gets the Excel Application object for the current Excel process.
     *
     * This can then be used to call back into Excel using the Excel
     * automation API, in the same way as VBA can be used to automate
     * Excel.
     *
     * The _Application object and objects obtained from it should only
     * be used from the same thread as it was obtained on. Sharing it
     * between threads will cause issues, and may cause Excel to crash.
     *
     * @param xl The ExcelAddIn object obtained from Jinx.
     * @return An Excel Application instance.
     */
    public static _Application getApplication(ExcelAddIn xl) {
        Com4jObject unk = COM4J.wrapSta(Com4jObject.class, xl.getExcelApplication());
        return unk.queryInterface(_Application.class);
    }
}
