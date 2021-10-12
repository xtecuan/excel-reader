package com.xtesoft.samples.excelreader.resources;

import com.xtesoft.samples.excelreader.entities.Persona;
import com.xtesoft.samples.excelreader.services.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("/excel")
public class ExcelResource {

    @Autowired
    private ExcelService excelService;

    @RequestMapping(method = RequestMethod.GET)
    private List<Persona> readExcel(){
        return excelService.readXLSXCalificacionFile1();
    }

    @RequestMapping(method = RequestMethod.POST)
    private void save(){
        excelService.saveToDatabase();
    }
}
