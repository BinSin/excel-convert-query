package com.example.convert.controller;

import com.example.convert.dto.InsertHistoryQueryDto;
import com.example.convert.dto.InsertQueryDto;
import com.example.convert.dto.UpdateQueryDto;
import com.example.convert.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping(value = "/excel/convert/query")
@RequiredArgsConstructor
public class ExcelController {

  private final ExcelService excelService;

  @PostMapping(value = "/insert", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE,
      MediaType.APPLICATION_JSON_VALUE}, produces = MediaType.TEXT_PLAIN_VALUE)
  public ResponseEntity<String> convertToInsertQuery(
      @RequestPart("insertQueryDto") InsertQueryDto insertQueryDto,
      @RequestPart("file") MultipartFile file) throws Exception {
    return ResponseEntity.ok(excelService.convertInsertQuery(insertQueryDto, file));
  }

  @PostMapping(value = "/update", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE,
      MediaType.APPLICATION_JSON_VALUE}, produces = MediaType.TEXT_PLAIN_VALUE)
  public ResponseEntity<String> convertToUpdateQuery(
      @RequestPart("updateQueryDto") UpdateQueryDto updateQueryDto,
      @RequestPart("file") MultipartFile file) throws Exception {
    return ResponseEntity.ok(excelService.convertUpdateQuery(updateQueryDto, file));
  }

  @PostMapping(value = "/insert-history", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE,
      MediaType.APPLICATION_JSON_VALUE}, produces = MediaType.TEXT_PLAIN_VALUE)
  public ResponseEntity<String> convertToInsertHistoryQuery(
      @RequestPart("InsertHistoryQueryDto") InsertHistoryQueryDto insertHistoryQueryDto,
      @RequestPart("file") MultipartFile file) throws Exception {
    return ResponseEntity.ok(excelService.convertInsertHistoryQuery(insertHistoryQueryDto, file));
  }

}
