#!/bin/sh
PROJECT_DIR=$1
CONFIGNAME=$2
SEDEXCEL_DIR=$HOME/private/Projects/sedexcel
for file in ${PROJECT_DIR}/*.xlsx;
do
    fromexcelfile=${file}
    toexcelfile=`echo ${file} | sed 's/.xlsx/_new.xlsx/'`
    configfile=$CONFIGNAME
    java -jar $SEDEXCEL_DIR/target/scala-2.11/sedexcel-assembly-1.0.jar $fromexcelfile $toexcelfile $configfile
done
