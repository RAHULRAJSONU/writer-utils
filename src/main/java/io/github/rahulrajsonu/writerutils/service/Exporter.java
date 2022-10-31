package io.github.rahulrajsonu.writerutils.service;

import java.util.List;

public abstract class Exporter {

    private String exportType;

    public Exporter(){}

    public Exporter(String exportType){
        this.exportType = exportType;
    }
    public final <T> byte[] export(List<T> data, String sheetName){
        Writer writer = createWriter(exportType);
        return writer.write(data,sheetName);
    }

    public abstract Writer createWriter(String exportType);

}
