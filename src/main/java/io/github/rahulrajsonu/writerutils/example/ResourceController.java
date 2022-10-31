package io.github.rahulrajsonu.writerutils.example;

import io.github.rahulrajsonu.writerutils.factory.SpreadsheetExporter;
import io.github.rahulrajsonu.writerutils.service.Exporter;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@RestController
@RequestMapping("/download")
public class ResourceController {

    @PostMapping()
    public ResponseEntity<StreamingResponseBody> downloadFile(@RequestBody List<XlsxUser> user,
                                                              HttpServletResponse response) {
        response.setHeader("Content-Disposition", "attachment; filename=User.xls");
        int BUFFER_SIZE = 1024;

        Exporter exporter = new SpreadsheetExporter();

        return ResponseEntity.ok(outputStream -> {
            int bytesRead;
            byte[] buffer = new byte[BUFFER_SIZE];

            //convert the file to InputStream
            InputStream inputStream = new ByteArrayInputStream(exporter.export(user,XlsxUser.class,"User.xls"));
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
                outputStream.flush();
            }
        });

    }

    @GetMapping("/getPI")
    public ResponseEntity<List<XlsxUser>> getPersonalInfo(){
        List<XlsxUser> data = new ArrayList<>();
        for(int i=0; i < 10; i++){
            data.add(new XlsxUser("Person: "+i,i%2==0?"Female":"Male",28,22.7,new Date(1653868800000L),false, Collections.emptyList(), Collections.emptyList()));
        }
        return ResponseEntity.ok(data);
    }
}
