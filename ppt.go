package main

import (
    "fmt"
    "github.com/xuri/excelize"
    "encoding/csv"
    "strings"
    "io/ioutil"
    //"sort"
    "strconv"
)
var xlsxWrite *excelize.File
var xlsxRead1 *excelize.File
var xlsxRead2 []byte
var err error
var lenrows int
var procsite int
var paindatas []string
var producttables []string
var firstName string
var readtable []string
var ss[][]string
var sz int
var addr1 string ="C:/Users/EDZ/Desktop/CHPA市场匹配表/市场定义表/Market+Definition+for+Report+Automation+1108V3.xlsx"
var addr2 string
var insertRow int =2

//var thelastindex int=2
func main() {
    ReadTable()
    readtable= RemoveDuplicatesAndEmpty(readtable)
    
    for _,data:=range readtable{
       SetFormat()
        insertRow=2
        for i:=2;i<13;i++{
            Place:="B"+strconv.Itoa(i)
            cellB := xlsxRead1.GetCellValue("New product Market definition", Place)
            if len(cellB)>0&&data==cellB[0:len(data)]{
                Place="E"+strconv.Itoa(i)
                cellE := xlsxRead1.GetCellValue("New product Market definition", Place)[0:5]
                if cellE!="See b"{
                    Place="C"+strconv.Itoa(i)
                    cellC:= xlsxRead1.GetCellValue("New product Market definition", Place)
                    datas:=strings.SplitAfter(cellC, "\n") 
                    if data=="Oncology Elunate"{
                        WriteFirstOnoE(data,cellB,datas)
                        RepeatFirstE(data,cellB,datas)
                    }else if data=="Oncology Tyvyt"{
                        if cellB=="Oncology Tyvyt Oncology market"{
                            WriteFirstOnoT(data,cellB,datas)
                        }else if cellB=="Oncology Tyvyt PD-1 market"{
                            WriteFirstOnoE(data,cellB,datas)
                            RepeatFirstE(data,cellB,datas)
                        }
                    }else{
                        WriteFirst(data,cellB,datas)
                        RepeatFirst(data,cellB,datas)
                        //ReadPainDates(data,datas)   
                        //WriteSecond(data)
                    }   
                }else if cellE=="See b"{
                    Place="C"+strconv.Itoa(i)
                    cellC:= xlsxRead1.GetCellValue("New product Market definition", Place)
                    datas:=strings.SplitAfter(cellC, "\n") 
                    WriteFirst(data,cellB,datas)                   
                    //ReadPainDates()
                   // Save(data)
                    ReadPainDates(data,datas)   
                    WriteSecond(data)
                    break
                }
            } 
        }
        Save(data)
    } 
    //Read(addr1)
    //ReadPainDates()   
    //paindatas= RemoveDuplicatesAndEmpty(paindatas)
    //SetFormat()
    //WriteFirst()
    //WriteSecond()  
}
func WriteFirstOnoT(str string,strb string,datas[]string){
    addr2="C:/Users/EDZ/Desktop/CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    //fmt.Println(ss)
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][0]+"\n"
            /*if i==458{
                fmt.Println("aaaaaaaaaaaaaa")
            }*/
            if data==data2{
                SetRow(strb,
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING")
            }
         }
    }
}
func WriteFirstOnoE(str string,strb string,datas[]string){
    addr2="C:/Users/EDZ/Desktop/CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    //fmt.Println(ss)
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][7]+"\n"
            /*if i==458{
                fmt.Println("aaaaaaaaaaaaaa")
            }*/
            if data==data2{
                SetRow(strb,
                data2,
                ss[i][9],
                ss[i][13],
                strb+"MAPPING")
            }
         }
    }
}
func RepeatFirstE(str string,strb string,datas[]string){
    addr2="C:/Users/EDZ/Desktop/CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    //fmt.Println(ss)
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][7]+"\n"
           /* if i==458{
                fmt.Println("aaaaaaaaaaaaaa")
            }*/
            if data==data2{
                SetRow(ss[i][9],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING")
            }
         }
    }
}

func ReadTable(){
    xlsxRead1, err = excelize.OpenFile(addr1)
    if err != nil {
        fmt.Println(err)
        return
    }
    rows:=xlsxRead1.GetRows("New product Market definition")
    lenrows=len(rows)
    s := make([]string,13)
    for i:=2;i<=11;i++{
        Place:="A"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue("New product Market definition", Place)
        if len(cell)>0{
            s[i-2]=cell
        }
    }
    readtable=s
}
func SetFormat(){
    
    xlsxWrite=excelize.NewFile()
    index := xlsxWrite.NewSheet("Sheet1")
     //SetRow("DISPLAY","COMPS DESC","PRODUCT DESC","PACK DESC","Name")
     xlsxWrite.SetCellValue("Sheet1", "A1", "DISPLAY")
     xlsxWrite.SetCellValue("Sheet1", "B1", "COMPS DESC")
     xlsxWrite.SetCellValue("Sheet1", "C1", "PRODUCT DESC")
     xlsxWrite.SetCellValue("Sheet1", "D1", "PACK DESC")
     xlsxWrite.SetCellValue("Sheet1", "E1", "Name")
     xlsxWrite.SetActiveSheet(index) 
	/*xlsxWrite:= excelize.NewFile()
    index := xlsxWrite.NewSheet("Sheet1")
    SetRow("Display","COMPS DESC","PRODUCT DESC","PACK DESC","Name")
	/*xlsxWrite.SetCellValue("Sheet1", "A1", "Display")
    xlsxWrite.SetCellValue("Sheet1", "B1", "COMPS DESC")
    xlsxWrite.SetCellValue("Sheet1", "C1", "PRODUCT DESC")
	xlsxWrite.SetCellValue("Sheet1", "D1", "PACK DESC")
	xlsxWrite.SetCellValue("Sheet1", "E1", "Name")
    xlsxWrite.SetActiveSheet(index) */
}
func Save(str string){
	err := xlsxWrite.SaveAs(str+"_CHPA_ BRAND_1_MAPPING.xlsx")
    if err != nil {
        fmt.Println(err)
    }
}
/*func Read(str string){//读第一个文件
    xlsxRead1, err = excelize.OpenFile(str)
    if err != nil {
        fmt.Println(err)
        return
    }
    //cell := xlsxRead1.GetCellValue("Sheet1", "B2")
    //fmt.Println(cell)

    firstName=xlsxRead1.GetCellValue("New product Market definition", "B2")
    Place:=FindMolecule()
    cell := xlsxRead1.GetCellValue("New product Market definition", Place)
    datas=strings.SplitAfter(cell, "\n") 
    //sort.Strings(datas)
    //fmt.Println(RemoveDuplicatesAndEmpty(datas))

}*/
func ReadPainDates(str string,datas []string){
    s := make([]string,50)
    for i:=1;i<=lenrows;i++{
        Place:="A"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue("New product Market definition", Place)      
        strformat:=str+"显示格式"
        if cell==strformat{
            i+=2
            tmp:=i
            Place="A"+strconv.Itoa(i)
            cell = xlsxRead1.GetCellValue("New product Market definition", Place)
            strproc:=str+"产品对应关系如下："
           for cell!=strproc{    
               s[i-tmp]=cell
               i++
               site:="A"+strconv.Itoa(i)
               cell=xlsxRead1.GetCellValue("New product Market definition", site)        
           }
           procsite=i
           paindatas=s
           return
        }
    }
}
func FindMolecule() string {
    rows := xlsxRead1.GetRows("New product Market definition")
    lenrows=len(rows)
    for i:=1;i<=lenrows;i++{
        Place:="D"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue("New product Market definition", Place)
        if cell=="Molecule"{
            return "C"+strconv.Itoa(i)
        }
    }
    //xlsx.GetCellValue("Sheet1", "B2")
    return ""
}
func RemoveDuplicatesAndEmpty(a []string ) (ret [] string){
    a_len := len(a)
    for i:=0; i < a_len; i++{
        if (i > 0 && a[i-1] == a[i]) || len(a[i])==0{
            continue;
        }
        ret = append(ret, a[i])
    }
    return
}
func SetRow(strs...string){
    
    xlsxWrite.SetCellValue("Sheet1", "A"+strconv.Itoa(insertRow), strs[0])
    xlsxWrite.SetCellValue("Sheet1", "B"+strconv.Itoa(insertRow),strs[1])
    xlsxWrite.SetCellValue("Sheet1", "C"+strconv.Itoa(insertRow), strs[2])
	xlsxWrite.SetCellValue("Sheet1", "D"+strconv.Itoa(insertRow), strs[3])
    xlsxWrite.SetCellValue("Sheet1", "E"+strconv.Itoa(insertRow),strs[4])
    insertRow++
}

func WriteFirst(str string,strb string,datas[]string){
    addr2="C:/Users/EDZ/Desktop/CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    //fmt.Println(ss)
    sz = len(ss)
    for _,data:=range datas{
        for i:=0;i<sz;i++{
            data2:=ss[i][1]+"\n"
            if data==data2{
                SetRow(strb,
                ss[i][1],
                ss[i][3],
                ss[i][5],
                strb+"MAPPING")
            }
         }
    }
}

func RepeatFirst(str string,strb string,datas[]string){
    addr2="C:/Users/EDZ/Desktop/CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    //fmt.Println(ss)
    sz = len(ss)
    for _,data:=range datas{
        for i:=0;i<sz;i++{
            data2:=ss[i][1]+"\n"
            if data==data2{
                SetRow(data,
                ss[i][1],
                ss[i][3],
                ss[i][5],
                strb+"MAPPING")
            }
         }
    }
}
func WriteSecond(str string){
    //tmptotaldatas := make([]string,len(paindatas))
    //producttablesite:=GetProductSite(str)
    producttablesite:=procsite
    if str=="Cymbalta CMP"{
        flagtotal:=0
        flagall:=0
        for i:=0;i<len(paindatas);i++{
            if len(paindatas[i])>0{
                totalhead:=paindatas[i][0:5]
                allhead:=paindatas[i][0:9]
                if totalhead=="Total"&&flagtotal==0{
                    // name:=paindatas[i][6:]
                    flagtotal=1
                    for j:=producttablesite;j<=lenrows;j++{                       
                        Place:="A"+strconv.Itoa(j)
                        cell1 := xlsxRead1.GetCellValue("New product Market definition", Place)
                        Place ="C"+strconv.Itoa(j)
                        cell2 := strings.SplitAfter(xlsxRead1.GetCellValue("New product Market definition", Place)," ")[0]
                        if cell1=="NSAIDs"&&(cell2=="Other "||cell2=="All "){ 
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("Total NSAIDs",cell3)
                        }else if cell1 == "Weak Opioids"&&(cell2=="Other "||cell2=="All "){
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("Total Weak Opioids",cell3)
                        }else if cell1 == "Strong Opioids"&&(cell2=="Other "||cell2=="All "){
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("Total Strong Opioids",cell3)
                        }else if cell1 == "Muscle Relaxant"&&(cell2=="Other "||cell2=="All "){
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("Total Muscle Relaxant",cell3)
                        }
                    }        
                }else if allhead =="All Other"&&flagall==0{
                    flagall=1
                    for j:=producttablesite;j<=lenrows;j++{                       
                        Place:="A"+strconv.Itoa(j)
                        cell1 := xlsxRead1.GetCellValue("New product Market definition", Place)               
                        if cell1=="NSAIDs"{ 
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("All Other NSAIDs",cell3)
                        }else if cell1 == "Weak Opioids"{
                            Place2:="B"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2("All Other Weak Opioids",cell3)
                        }
                    }
                }else {
                    for j:=producttablesite;j<=lenrows;j++{
                        Place:="C"+strconv.Itoa(j)
                        cell := xlsxRead1.GetCellValue("New product Market definition", Place)               
                        if cell==paindatas[i]{
                            Place2:="B"+strconv.Itoa(j)
                            cell2 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc(cell,cell2)
                        }
                    }
                    fmt.Println("over")
                }
            }
        }
    }else if str=="Cialis BPH"{
        //flagtotal:=0
        for i:=0;i<len(paindatas);i++{
            if len(paindatas[i])>0{
                totalhead:=paindatas[i][0:4]
                //allhead:=paindatas[i][0:9]
                if (totalhead=="PDE5"||totalhead =="a-bl"||totalhead =="5ARI"){
                    // name:=paindatas[i][6:]
                   // flagtotal=1
                    for j:=producttablesite;j<=lenrows;j++{                       
                        /*Place:="A"+strconv.Itoa(j)
                        cell1 := xlsxRead1.GetCellValue("New product Market definition", Place)*/
                        Place:="E"+strconv.Itoa(j)
                        cell2 := xlsxRead1.GetCellValue("New product Market definition", Place)
                        if cell2==paindatas[i]{
                            Place2:="A"+strconv.Itoa(j)
                            cell3 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc2(cell2,cell3)
                        }
                    }        
                }else {
                    for j:=producttablesite;j<=lenrows;j++{
                        Place:="D"+strconv.Itoa(j)
                        cell := xlsxRead1.GetCellValue("New product Market definition", Place)               
                        if cell==paindatas[i]{
                            Place2:="A"+strconv.Itoa(j)
                            cell2 := xlsxRead1.GetCellValue("New product Market definition", Place2)
                            WriteCompsDesc(cell,cell2)
                        }
                    }
                    fmt.Println("over")
                }
            }
        }
    }

}

func GetProductSite(str string) int {
    for i:=1;i<=lenrows;i++{
        Place:="A"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue("New product Market definition", Place)
        //s := make([]string,lenrows-i)
        if cell=="Category"{
            i++
            return i
            /*tmp:=i
            Place="C"+strconv.Itoa(i)
            cell = xlsxRead1.GetCellValue("New product Market definition", Place)
            for cell!=""{    
                s[i-tmp]=cell
                i++
                site:="C"+strconv.Itoa(i)
                cell=xlsxRead1.GetCellValue("New product Market definition", site)        
            }
            producttables=s
            return  */  
        }
    }
    return 0
}
func WriteCompsDesc(data1 string,data2 string){
    for i:=0;i<sz;i++{
        //data:=ss[i][1]
        if //data==data2||
        data1==ss[i][3] {     
            SetRow(data1,
            ss[i][1],
            ss[i][3],
            ss[i][5],
            "Cymbalta CMP _CHPA_ BRAND_1_MAPPING")
        }
    } 
}
func WriteCompsDesc2(data1 string,data2 string){
    for i:=0;i<sz;i++{
        data:=ss[i][1]
        if data==data2{     
            SetRow(data1,
            ss[i][1],
            ss[i][3],
            ss[i][5],
            "Cymbalta CMP _CHPA_ BRAND_1_MAPPING")
        }
    } 
}