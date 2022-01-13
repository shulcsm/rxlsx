Python 3.9.5

openpyxl==3.0.9
PyExcelerate==0.10.0
XlsxWriter==3.0.2

```
$ hyperfine ./bench_openpyxl.py ./bench_pyexcelerate.py ./bench_xlsxwriter.py ./bench_rxlsx.py 
Benchmark 1: ./bench_openpyxl.py
  Time (mean ± σ):     12.865 s ±  1.151 s    [User: 12.558 s, System: 0.249 s]
  Range (min … max):   10.963 s … 13.942 s    10 runs
 
Benchmark 2: ./bench_pyexcelerate.py
  Time (mean ± σ):      6.907 s ±  0.136 s    [User: 6.876 s, System: 0.022 s]
  Range (min … max):    6.775 s …  7.137 s    10 runs
 
Benchmark 3: ./bench_xlsxwriter.py
  Time (mean ± σ):      9.118 s ±  0.077 s    [User: 8.903 s, System: 0.092 s]
  Range (min … max):    8.998 s …  9.250 s    10 runs
 
Benchmark 4: ./bench_rxlsx.py
  Time (mean ± σ):      2.423 s ±  0.072 s    [User: 2.384 s, System: 0.036 s]
  Range (min … max):    2.345 s …  2.589 s    10 runs
 
Summary
  './bench_rxlsx.py' ran
    2.85 ± 0.10 times faster than './bench_pyexcelerate.py'
    3.76 ± 0.12 times faster than './bench_xlsxwriter.py'
    5.31 ± 0.50 times faster than './bench_openpyxl.py'
```