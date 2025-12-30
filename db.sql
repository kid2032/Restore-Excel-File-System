CREATE DATABASE excelvc;

USE excelvc;

CREATE TABLE excel_files
  (
     id        INT auto_increment PRIMARY KEY,
     file_path VARCHAR(500) UNIQUE,
     file_name VARCHAR(255),
     last_hash CHAR(64)
  );

CREATE TABLE excel_versions
  (
     id             INT auto_increment PRIMARY KEY,
     file_id        INT,
     version_number INT,
     file_data      LONGBLOB,
     created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
     UNIQUE(file_id, version_number),
     FOREIGN KEY (file_id) REFERENCES excel_files(id)
  ); 
