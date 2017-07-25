##### PHPExcel demo
原计划基于PHPExcel对常见的操作进行封装 经考虑 提供demo更加方便

##### 关于读
推荐使用 PHPExcel_IOFactory::load()方法加载excel，然后进行读，测试xls、xlsx没有兼容问题

##### PHPExcel使用时遇到的一些问题
测试中发现使用PHPExcel_Writer_Excel5对xls类型的excel在保存后再次读时会出错
