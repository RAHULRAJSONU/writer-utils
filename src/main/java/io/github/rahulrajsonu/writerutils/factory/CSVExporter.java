package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.service.Exporter;
import io.github.rahulrajsonu.writerutils.service.Writer;

public class CSVExporter extends Exporter {

    public CSVExporter(String exportType){
        super(exportType);
    }
    @Override
    public Writer createWriter(String exportType) {
        return new CSVWriter();
    }
}
