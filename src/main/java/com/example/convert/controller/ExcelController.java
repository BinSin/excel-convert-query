package com.example.convert.controller;

import com.example.convert.dto.InsertQueryDto;
import com.example.convert.dto.UpdateQueryDto;
import com.example.convert.service.ExcelService;
import io.swagger.v3.oas.annotations.Parameter;
import lombok.RequiredArgsConstructor;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping(value = "/excel/convert/query")
@RequiredArgsConstructor
public class ExcelController {

  private final ExcelService excelService;

  @PostMapping(value = "/insert", consumes = { MediaType.MULTIPART_FORM_DATA_VALUE, MediaType.APPLICATION_JSON_VALUE }, produces = MediaType.TEXT_PLAIN_VALUE)
  public ResponseEntity<String> convertToInsertQuery(
          @RequestPart("insertQueryDto") InsertQueryDto insertQueryDto,
          @RequestPart("file") MultipartFile file) throws Exception {
    return ResponseEntity.ok(excelService.convertToQuery(insertQueryDto, file));
  }

  @PostMapping(value = "/update", consumes = { MediaType.MULTIPART_FORM_DATA_VALUE, MediaType.APPLICATION_JSON_VALUE }, produces = MediaType.TEXT_PLAIN_VALUE)
  public ResponseEntity<String> convertToUpdateQuery(
          @RequestPart("updateQueryDto") UpdateQueryDto updateQueryDto,
          @RequestPart("file") MultipartFile file) throws Exception {
    return ResponseEntity.ok(excelService.convertToQuery(updateQueryDto, file));
  }

}
