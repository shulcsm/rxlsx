Python 3.9.5

openpyxl==3.0.9
PyExcelerate==0.10.0
XlsxWriter==3.0.2

```
$ hyperfine ./bench_openpyxl.py ./bench_pyexcelerate.py ./bench_xlsxwriter.py ./bench_rxlsx.py
Benchmark 1: ./bench_openpyxl.py
  Time (mean ± σ):      4.621 s ±  0.097 s    [User: 4.508 s, System: 0.096 s]
  Range (min … max):    4.509 s …  4.836 s    10 runs
 
Benchmark 2: ./bench_pyexcelerate.py
  Time (mean ± σ):      2.397 s ±  0.038 s    [User: 2.381 s, System: 0.016 s]
  Range (min … max):    2.337 s …  2.455 s    10 runs
 
Benchmark 3: ./bench_xlsxwriter.py
  Time (mean ± σ):      3.066 s ±  0.055 s    [User: 2.987 s, System: 0.038 s]
  Range (min … max):    2.988 s …  3.148 s    10 runs
 
Benchmark 4: ./bench_rxlsx.py
  Time (mean ± σ):     279.4 ms ±   4.5 ms    [User: 265.0 ms, System: 14.5 ms]
  Range (min … max):   273.2 ms … 287.2 ms    10 runs
 
Summary
  './bench_rxlsx.py' ran
    8.58 ± 0.19 times faster than './bench_pyexcelerate.py'
   10.97 ± 0.26 times faster than './bench_xlsxwriter.py'
   16.54 ± 0.44 times faster than './bench_openpyxl.py'
```