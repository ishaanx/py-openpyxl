#!/bin/bash
split -l 20000 "input.csv" output-file-part-
for i in output-file-part-*;
  do
    mv "$i" "$i.csv";
  done
