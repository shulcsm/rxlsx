Python 3.9.5

openpyxl==3.0.9
PyExcelerate==0.10.0
XlsxWriter==3.0.2

```
$ hyperfine ./bench_openpyxl.py ./bench_pyexcelerate.py ./bench_xlsxwriter.py ./bench_rxlsx.py 
Benchmark 1: ./bench_openpyxl.py
  Time (mean ± σ):     29.701 s ±  0.353 s    [User: 29.040 s, System: 0.498 s]
  Range (min … max):   29.139 s … 30.181 s    10 runs
 
Benchmark 2: ./bench_pyexcelerate.py
  Time (mean ± σ):     15.170 s ±  0.193 s    [User: 15.094 s, System: 0.048 s]
  Range (min … max):   14.940 s … 15.475 s    10 runs
 
Benchmark 3: ./bench_xlsxwriter.py
  Time (mean ± σ):     16.903 s ±  0.305 s    [User: 16.509 s, System: 0.147 s]
  Range (min … max):   16.425 s … 17.583 s    10 runs
 
Benchmark 4: ./bench_rxlsx.py
  Time (mean ± σ):      2.417 s ±  0.048 s    [User: 2.373 s, System: 0.037 s]
  Range (min … max):    2.362 s …  2.498 s    10 runs
 
Summary
  './bench_rxlsx.py' ran
    6.28 ± 0.15 times faster than './bench_pyexcelerate.py'
    6.99 ± 0.19 times faster than './bench_xlsxwriter.py'
   12.29 ± 0.28 times faster than './bench_openpyxl.py'
```