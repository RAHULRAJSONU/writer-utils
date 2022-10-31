package io.github.rahulrajsonu.writerutils.service;

import java.io.ByteArrayOutputStream;
import java.util.List;

public abstract class Exporter {
    public final <T> byte[] export(List<T> data, Class<T> clazz, String sheetName){
        Writer writer = createWriter();
        return writer.write(data, clazz,sheetName);
    }

    public abstract Writer createWriter();
}
