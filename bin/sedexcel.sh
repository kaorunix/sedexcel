#!/bin/sh
if [ $# -ne 1 ];then
    echo "Usage $0 EXCELFILE"
    exit 1
else
#	PROJECT_DIR=$1
    EXCELFILE=$1
    echo `pwd` | grep -q "10_見積"; ret=$?
#    if [ $ret -eq 1 ]; then
#	echo "You should execute in the directory 10_見積"
#	exit 1
#    fi
    SEDEXCEL_DIR=$HOME/private/Projects/sedexcel

    for file in *AOKI*.conf;
    do
	fromexcelfile=`echo ${EXCELFILE} | sed 's/.xlsx/_old.xlsx/'`
	toexcelfile=${EXCELFILE}
	configfile="../30_config/"`echo ${EXCELFILE} | sed 's/.xlsx/.conf/'`
	source $file
	echo "HDID=$HDID"
	echo "PJNAME=$PJNAME"
	echo "PRICE=$PRICE"
	echo "SUBMIT_DATE=$SUBMIT_DATE"
	cat $configfile \
	    | sed -e "/#HDID_START/,/#HDID_END/ s/\([^=]*=\).*/\1${HDID}/" \
	    | sed -e "/#PJNAME_START/,/#PJNAME_END/ s/\([^=]*=\).*/\1${PJNAME}/" \
	    | sed -e "/#PRICE_START/,/#PRICE_END/ s/\([^=]*=\).*/\1${PRICE}/" \
	    | sed -e "/#SUBMIT_DATE_START/,/#SUBMIT_DATE_END/ s!\([^=]*=\).*!\1${SUBMIT_DATE}!" > ${configfile}.tmp
		  
	mv ${configfile}.tmp $configfile
	mv $toexcelfile $fromexcelfile
	java -jar $SEDEXCEL_DIR/target/scala-2.11/sedexcel-assembly-1.0.jar $fromexcelfile $toexcelfile $configfile
    done
fi
