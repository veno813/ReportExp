package main

import (
	"fmt"
	"time"
)

type Fundinfo struct {
	FundCode     string //基金代码
	FundName     string //基金名称
	FundFullName string //基金全称
}

var filepath string = "Z:/003 注册登记 TA/003 业务报表报送/001 中证登持有人份额报送/0-中证登报表拆分/file/"

//程序功能：
// 1.拆分中保登报表，将全量报表拆分为各个基金的报表
// 2.提供导出进度条

/*程序流程:
1.读取sheet2中的导出基金列表
2.根据导出基金列表中的值匹配sheet1中的全量数据
3.需要注意，将10亿以上数字的小数点位数删除
4.将匹配数据根据模板文件格式输出，文件名为基金代码基金名称（例：999214月月盈）
*/
func main() {
	fmt.Println("********************************************************************")
	fmt.Println("********************************************************************")
	fmt.Println("************************中登季度报表导出工具************************")
	fmt.Println("********************************************************************")
	fmt.Println("********************************************************************")

	fmt.Println("请输入全量文件路径：")
	fmt.Println("默认地址为：" + filepath)
	var s string
	fmt.Scanln(&s)
	if s == "" {
		t1 := time.Now()
		fmt.Println("开始拆分数据", t1)
		fetchDataByExcel()
		elapsed := time.Since(t1)
		fmt.Println("导出耗时:", elapsed)
	}
	fmt.Scanln(&s)

}
