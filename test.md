# download

```go

package main

import (
	"log"
    "fmt"
	"io/ioutil"

    "github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

var colStrs= []string{"A", "B","C","D","E","F","G","H","I","J","K","L","M","N","O","P"}
var titles=[]string{"日期",	"按成交金额排名",	"按访客指数排名",	"访客指数",	"搜索点击人气",	"搜索点击指数",	"关注人数",	"加购人数",	"成交金额指数",	"成交单量指数",	"成交转化率",	"成交客单价",	"品牌店铺数",	"动销店铺数",	"动销商品数",	"动销商品占比"}


var y=2  

func main() {
	dirs, err := ioutil.ReadDir("./")
	if err != nil {
        log.Fatal(err)
    }
	for _,fi:=range dirs{
		if fi.IsDir(){
			y=2
			fmt.Println(fi.Name())
			merge(fi.Name())
		} 
	}
}

func merge(name string){ 
	xf := excelize.NewFile()
	dirs, err := ioutil.ReadDir(name)
	if err != nil {
        log.Fatal(err)
    }
	for i, title:=range titles{
		k:= fmt.Sprintf("%s%d",colStrs[i],1) 
		xf.SetCellValue("Sheet1", k, title)
	}
	for _,fi:=range dirs{
		fmt.Println(fi.Name())
		mergeItem(name+"/"+fi.Name(),xf)
	}
	if err := xf.SaveAs(name+".xlsx"); err != nil {
        fmt.Println(err)
    }
}

func mergeItem(xlsPath string,xf *excelize.File){
	xlsFile, err := xls.Open(xlsPath, "utf-8")
    if err != nil {
        log.Fatal(err)
    }
	sheet := xlsFile.GetSheet(0) 
	for j := 1; j < int(sheet.MaxRow)+1; j++ {
        xlsRow := sheet.Row(j)
        rowColCount := xlsRow.LastCol()
		for i := 0; i < rowColCount; i++ {
			k:= fmt.Sprintf("%s%d",colStrs[i],y) 
			xf.SetCellValue("Sheet1", k, xlsRow.Col(i))
		}
		y++
    }
}

```
