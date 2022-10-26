package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.service.Exporter;
import io.github.rahulrajsonu.writerutils.service.Writer;

public class CSVExporter extends Exporter {
    @Override
    public Writer createWriter() {
        return new CSVWriter();
    }
}
