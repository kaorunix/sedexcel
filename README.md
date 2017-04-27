# SedExcel

This program can replace string in excel file.

## How to use

Make config file, connected key and value with equal(=). Omit space between equal. You can use String, Number, Date as key.

### key

* Date you can assign like `1999/99/99`
* Number you can assign like `999999`
* String you can assign like `name`

### value

* Date you can assign like `2017/11/22`
* Number you can assign like `-100.01`
* String you can assign any words. You can use UTF-8.

### config example
```
1999/99/99=2017/11/22
999999=-100.01
name=鈴木
```

### make excel file
Create a excel file embedding words which put as key in config file. SedExcel will replace from words which are same charactors as key, to value. You can same words in the excel file, all of same words will replace same values.

