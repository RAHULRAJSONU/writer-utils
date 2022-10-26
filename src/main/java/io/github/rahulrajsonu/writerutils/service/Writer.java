package io.github.rahulrajsonu.writerutils.service;

import java.util.List;

public interface Writer {
    <T extends Exportable> byte[] write(List<T> data, Class<T> clazz, String fileName);
}
