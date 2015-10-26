#说明

修改来自[Excel格式导入导出数据（单语言版本）](http://www.mycncart.com/index.php?route=product/product&product_id=35)，感谢[原作者](http://www.mycncart.com/index.php?route=product/extension&nickname=%E6%9D%A8%E5%85%86%E9%94%8B)。

#Excel导入导出工具

用于Opencart，可以导入导出商品，导出订单地理分布（简单），销售报表（简单）。

使用了PHPExcel类。

#安装

确保Opencart已经安装vamod。

安装xml文件，然后刷新并为管理员组添加权限。可以在Tool/Excel Import Export Tool下找到。

#vqmod安装

下载[vqmod for opencart](http://www.opencart.com/index.php?route=extension/extension/info&extension_id=19501)，[vqmod](https://github.com/vqmod/vqmod)

把vqmod文件放到system同级的目录下，执行example.com/vqmod/install，注意权限；

把vqmod for opencart2 的文件解压替换。

注意：模块涉及到数据库中的有些表或字段已经更改，一般会出问题：)