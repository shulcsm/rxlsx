Python 3.9.5

openpyxl==3.0.9
PyExcelerate==0.10.0
XlsxWriter==3.0.2

```
$ hyperfine ./bench_openpyxl.py ./bench_pyexcelerate.py ./bench_xlsxwriter.py ./bench_rxlsx.py 
Benchmark 1: ./bench_openpyxl.py
  Time (mean ± σ):     30.275 s ±  0.322 s    [User: 29.587 s, System: 0.518 s]
  Range (min … max):   29.537 s … 30.570 s    10 runs
 
Benchmark 2: ./bench_pyexcelerate.py
  Time (mean ± σ):     15.163 s ±  0.272 s    [User: 15.113 s, System: 0.027 s]
  Range (min … max):   14.581 s … 15.468 s    10 runs
 
Benchmark 3: ./bench_xlsxwriter.py
  Time (mean ± σ):     17.060 s ±  0.277 s    [User: 16.719 s, System: 0.163 s]
  Range (min … max):   16.646 s … 17.515 s    10 runs
 
Benchmark 4: ./bench_rxlsx.py
  Time (mean ± σ):      2.690 s ±  0.048 s    [User: 2.650 s, System: 0.035 s]
  Range (min … max):    2.623 s …  2.778 s    10 runs
 
Summary
  './bench_rxlsx.py' ran
    5.64 ± 0.14 times faster than './bench_pyexcelerate.py'
    6.34 ± 0.15 times faster than './bench_xlsxwriter.py'
   11.25 ± 0.23 times faster than './bench_openpyxl.py'
```