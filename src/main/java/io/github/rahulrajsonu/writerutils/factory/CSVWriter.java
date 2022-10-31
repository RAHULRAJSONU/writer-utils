package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.service.Writer;

import java.util.List;

public class CSVWriter implements Writer {

    @Override
    public <T> byte[] write(List<T> data, String fileName) {
        return null;
    }
}
