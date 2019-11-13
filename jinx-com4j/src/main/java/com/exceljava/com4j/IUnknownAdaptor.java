package com.exceljava.com4j;

import com.exceljava.jinx.IUnknown;
import com4j.Com4jObject;

/**
 * Adaptor class that implements the IUnknown interface
 * wrapping a Com4jObject instance.
 */
class IUnknownAdaptor implements IUnknown {
    private Com4jObject obj;

    public IUnknownAdaptor(Com4jObject obj) {
        this.obj = obj;
    }

    private static <T extends Com4jObject> Class<T> adaptClass(Class<?> cls) {
        return (Class<T>)cls;
    }

    @Override
    public <T> T queryInterface(Class<T> cls) {
        Com4jObject obj = this.obj;
        if (null != obj && cls.isAssignableFrom(Com4jObject.class)) {
            return obj.queryInterface(adaptClass(cls));
        }
        return null;
    }

    @Override
    public long getPointer(boolean addRef) {
        if (addRef) {
            throw new UnsupportedOperationException();
        }

        Com4jObject obj = this.obj;
        if (null != obj) {
            return obj.getIUnknownPointer();
        }

        return 0;
    }

    @Override
    public void close() {
        this.obj = null;
    }
}
