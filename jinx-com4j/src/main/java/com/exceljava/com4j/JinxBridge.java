package com.exceljava.com4j;

import com.exceljava.jinx.ExcelAddIn;
import com.exceljava.jinx.ExcelArgumentConverter;
import com.exceljava.jinx.IUnknown;
import com.exceljava.com4j.excel._Application;
import com4j.COM4J;
import com4j.Com4jObject;


/**
 * Bridge between the Jinx add-in and com4j.
 * Used for obtaining com4j COM wrappers from code running in Excel using Jinx.
 */
public class JinxBridge {
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
        return xl.getExcelApplication(_Application.class);
    }

    /**
     * Converts IUnknown to any Com4jObject type.
     *
     * This allows Com4jObjects to be used as method arguments where
     * an IUnknown would be passed by Jinx. For example, in the
     * ribbon actions.
     *
     * @param unk IUnknown instance.
     * @param cls Class of type to cast to.
     * @param <T> Type to cast to.
     * @return IUnknown instance cast to a Com4jObject instance.
     */
    @ExcelArgumentConverter
    public static <T extends Com4jObject> T convertIUnknown(IUnknown unk, Class<T> cls) {
        Com4jObject obj = COM4J.wrapSta(Com4jObject.class, unk.getPointer(true));
        return obj.queryInterface(cls);
    }
}
