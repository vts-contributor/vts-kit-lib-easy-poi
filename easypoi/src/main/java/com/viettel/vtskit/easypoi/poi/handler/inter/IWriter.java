package com.viettel.vtskit.easypoi.poi.handler.inter;

import java.util.Collection;
/**
 * Big data writes out the service interface
 *
 * @author liusq
 */
public interface IWriter<T> {
    /**
     *
     * @return
     */
    default public T get() {
        return null;
    }

    /**
     *
     * @return
     */
    public IWriter<T> write(Collection data);

    /**
     * Close the flow and become a business
     *
     * @return
     */
    public T close();
}
