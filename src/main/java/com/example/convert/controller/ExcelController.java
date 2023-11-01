package com.example.convert.controller;

import com.example.convert.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelController {

  private final ExcelService excelService;

  @PostMapping("/convert/insert-query")
  public String convertToQuery(@RequestParam("file") MultipartFile file)
      throws Exception {
    return excelService.convertToQuery(file);
  }

}
