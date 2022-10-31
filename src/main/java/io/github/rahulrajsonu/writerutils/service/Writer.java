package io.github.rahulrajsonu.writerutils.service;

import java.util.List;

public interface Writer {
    <T> byte[] write(List<T> data, String fileName);
}
