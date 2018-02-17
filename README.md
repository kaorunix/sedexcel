# SedExcel

This program can replace string in excel file.

## How to use

Make config file, connected key and value with equal(=). Omit space between equal.


### key

You should assign cell position like `AB132`
This cell will replace the value of right side.

### value

You can use String, Number, Date as value.
The value will be treated as string.

### config example

```
BA15=2017/11/22
BB15=-100.01
D3=鈴木
```

### make excel file

1. Create a excel file as templete with xlsx as extention.
2. Create a config file.
3. Execute sedexcel.sh like follows.

```
$ sedexcel.sh TEMPLATE.xlsx ABC.conf
```

##### result

The original excel file will rename TEMPLATE_old.xlsx. And original name file will be created, the cells, which were configured in config file, will be embeded values.
