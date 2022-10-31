package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.service.Exportable;
import io.github.rahulrajsonu.writerutils.service.Writer;

import java.util.List;

public class CSVWriter implements Writer {

    @Override
    public <T> byte[] write(List<T> data, Class<T> clazz, String fileName) {
        return null;
    }
}
