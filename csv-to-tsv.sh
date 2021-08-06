#!/bin/bash

for file in *.csv
do
  mv "$file" "${file%.csv}.tsv"
done
