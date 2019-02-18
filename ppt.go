package main

import (
    "fmt"
    "github.com/xuri/excelize"
    "encoding/csv"
    "strings"
    "io/ioutil"
    "os"
    "strconv"
)
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
var w * csv.Writer
var addr1 string ="CHPA市场匹配表/市场定义表/Market+Definition+for+Report+Automation+1108V3.xlsx"
var addr2 string

var sheet string="Existing product Market definit"
var lastmak string=" "
func main() {
    ReadTable()
    readtable= RemoveDuplicatesAndEmpty(readtable)
    for _,data:=range readtable{
        if data=="Ceclor Solid "||data=="Ceclor Liquid "||data=="Ceclor Total "{
            continue
        }
        f, err := os.Create(data+"_CHPA_BRAND_1_MAPPING.csv") //创建文件
        if err != nil {
            panic(err)
        }
	    defer f.Close()
	    f.WriteString("\xEF\xBB\xBF") // 写入UTF-8 BOM
        w = csv.NewWriter(f) //创建一个新的写入文件流
        w.Write([]string{"DISPLAY", "COMPS DESC", "PRODUCT DESC","PACK DESC","Name"})  
        for i:=2;i<300;i++{
            if i==96{
                fmt.Println("aaa")
            }
            Place:="B"+strconv.Itoa(i)
            cellB:=xlsxRead1.GetCellValue(sheet, Place)
            if len(cellB)==0{
                break
            }
            if lastmak==cellB{
                continue
            }
            lastmak=cellB
            if len(cellB)>0&&data==cellB[0:len(data)]{
                Place="C"+strconv.Itoa(i)
                cellC := xlsxRead1.GetCellValue(sheet, Place)
                datas:=strings.SplitAfter(cellC, "\n")
                if cellC[0:3]=="L01"{
                    RepeatWriteFirstOnoT(data,cellB[len(data)+1:],datas)
                    WriteFirstOnoT(data,"Total Mkt",datas)
                    WriteFirstOnoT(data,cellB[len(data)+1:],datas)   
                    break  
                }else if cellB=="Prozac Lilly relevant MKT "{
                    WriteAnti(data,"Total Mkt",datas)   
                    WriteAnti(data,cellB[len(data)+1:],datas)           
                    RepeatWriteJ(data,cellB[len(data)+1:],datas)           
                }else if cellB=="Prozac Prozac AD Branded Market"{
                    WriteJ(data,"Total Mkt",datas) 
                    WriteJ(data,cellB[len(data)+1:],datas)             
                    RepeatWriteJ(data,cellB[len(data)+1:],datas)     
                    break    
                }else if cellB=="Zyprexa Branded MKT "{
                    WriteJ(data,"Total Mkt",datas) 
                    RepeatWriteZ(data,cellB[len(data)+1:],datas)
                    WriteZ(data,cellB[len(data)+1:],datas)
                }else if cellB=="Strattera Relevant market "{
                    RepeatWriteJ2(data,cellB[len(data)+1:],datas)
                    WriteJ(data,cellB[len(data)+1:],datas)  
                    RepeatWriteJ2(data,data,datas)
                    break
                }else if cellB=="CIALIS PDE-5"{                    
                    cellC:= xlsxRead1.GetCellValue("New product Market definition", "C4")
                    datas:=strings.SplitAfter(cellC, "\n") 
                    WriteFirst("Cialis BPH","Cialis BPH market",datas)                   
                    ReadPainDates("Cialis BPH",datas) 
                    datas= RemoveDuplicatesAndEmpty(datas)  
                    WriteSecond("Cialis BPH")     
                }else if cellB=="EVISTA WOMEN'S HEALTH Market"{
                    RepeatWriteZ(data,cellB[len(data)+1:],datas) 
                    WriteZ(data,cellB[len(data)+1:],datas) 
                    RepeatWriteZ(data,"OP MKT",datas)
                    RepeatWriteB(data,cellB[len(data)+1:],datas) 
                    break    
                }else if cellB=="Trulicity GLP-1 relevant market (A10S)"{
                    WriteD(data,"Total Mkt",datas) 
                    WriteD(data,cellB[len(data)+1:],datas) 
                    RepeatWriteD(data,cellB[len(data)+1:],datas)                   
                    break    
                }else if cellB=="Insulin Total Lilly Insulin"{
                    WriteA10C_D(data,"Lilly Insulin MKT",datas)    
                    WriteA10C_D(data,cellB[len(data)+1:],datas)    
                }else if cellB=="Insulin Total Animal Insulin"{
                    WriteA10D(data,"Animal Insulin MKT",datas) 
                    WriteA10D(data,cellB[len(data)+1:],datas)          
                }else if cellB=="Insulin Total Human Insulin"{
                    WriteFirstHuman(data,"Human Insulin MKT",datas) 
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)     
                    RepeatWriteHuman(data,cellB[len(data)+1:],datas)      
                }else if cellB=="Insulin Total Mealtime Analog"{
                    WriteFirstHuman(data,"Mealtime Analog MKT",datas) 
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)     
                    RepeatWriteHuman(data,cellB[len(data)+1:],datas)   
                }else if cellB=="Insulin Mealtime analog"{
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)
                    RepeatWriteFamily(data,cellB[len(data)+1:],datas)
                }else if cellB=="Insulin Rapid analog"{
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)
                    RepeatWriteRapid(data,cellB[len(data)+1:],datas)
                }else if cellB=="Insulin Mixture analog"{
                    WriteFirstHuman(data,"Mixture analog MKT",datas) 
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)     
                    RepeatWriteHuman(data,cellB[len(data)+1:],datas)    
                }else if cellB=="Insulin Basal Analog"{
                    WriteFirstHuman(data,"Basal Analog MKT",datas) 
                    WriteFirstHuman(data,cellB[len(data)+1:],datas)     
                    RepeatWriteHuman(data,cellB[len(data)+1:],datas)
                }else if cellB=="Insulin Total Mealtime Analog Kwikpen market"{
                    WriteFirstKwikpen(data,"Analog Kwikpen MKT",datas)
                    WriteFirstKwikpen(data,cellB[len(data)+1:],datas)
                    RepeatWriteKwikpen(data,cellB[len(data)+1:],datas)   
                }else if cellB=="Insulin Total Humulin Kwikpen Market"{
                    WriteFirstKwikpen(data,"Humulin Kwikpen MKT",datas)
                    WriteFirstKwikpen(data,cellB[len(data)+1:],datas)
                    RepeatWriteKwikpen(data,cellB[len(data)+1:],datas)
                }else { 
                    WriteFirstOnoE(data,"Total Mkt",datas)
                    WriteFirstOnoE(data,cellB[len(data)+1:],datas)
                    RepeatFirstE(data,cellB[len(data)+1:],datas)
                    w.Flush()
                  
                }
            } 
        }
        w.Flush()
    } 
}
func RepeatWriteFamily(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][7])<len(data)-1{
                continue
            }
            data2:=ss[i][7][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                w.Write([]string{ss[i][7]+" Family",
                ss[i][2],
                ss[i][7],
                ss[i][11],
            strb+"MAPPING"}) 
            }
         }
    }
}
func RepeatWriteRapid(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][7])<len(data)-1{
                continue
            }
            data2:=ss[i][7][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                w.Write([]string{ss[i][7]+" Rapid",ss[i][2],ss[i][7],ss[i][11],strb+"MAPPING"})
            }
         }
    }
}
func RepeatWriteKwikpen(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        data=data[0:4]
        if data=="HUMA"{
            for i:=0;i<sz;i++{
                if len(ss[i][11])<len(data)-1{
                    continue
                }
                data2:=ss[i][11][0:len(data)]+"\n"
                /*a1:=len(data)
                a2:=len(data2)
                fmt.Println(a1,a2)*/
                if data==data2{
                    w.Write([]string{ss[i][7],
                        ss[i][0],
                        ss[i][7],
                        ss[i][11],
                    strb+"MAPPING"})
                }
             }
        }else if data=="NOVO"{
            for i:=0;i<sz;i++{
                if len(ss[i][7])<len(data)-1{
                    continue
                }
                data2:=ss[i][7][0:len(data)]+"\n"
                /*a1:=len(data)
                a2:=len(data2)
                fmt.Println(a1,a2)*/
                if data==data2{
                    w.Write([]string{ss[i][7],
                        ss[i][2],
                        ss[i][7],
                        ss[i][11],
                    strb+"MAPPING"})
                }
             }
        }
       
    }
}
func WriteFirstKwikpen(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        data=data[0:4]
        if data=="HUMA"{
            for i:=0;i<sz;i++{
                if len(ss[i][11])<len(data)-1{
                    continue
                }
                data2:=ss[i][11][0:len(data)]
                /*a1:=len(data)
                a2:=len(data2)
                fmt.Println(a1,a2)*/
                if data==data2{
                     w.Write([]string{strb,
                        ss[i][2],
                        ss[i][7],
                        ss[i][11],
                    strb+"MAPPING"})
                }
             }
        }else if data=="NOVO"{
            for i:=0;i<sz;i++{
                if len(ss[i][7])<len(data)-1{
                    continue
                }
                data2:=ss[i][7][0:len(data)]
                /*a1:=len(data)
                a2:=len(data2)
                fmt.Println(a1,a2)*/
                if data==data2{
                     w.Write([]string{strb,
                        ss[i][0],
                        ss[i][7],
                        ss[i][11],
                    strb+"MAPPING"})
                }
             }
        }
       
    }
}
func WriteFirstHuman(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][7])<len(data)-1{
                continue
            }
            data2:=ss[i][7][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{strb,
                    ss[i][2],
                    ss[i][7],
                    ss[i][11],
                strb+"MAPPING"})
            }
         }
    }
}
func RepeatWriteHuman(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][7])<len(data)-1{
                continue
            }
            data2:=ss[i][7][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{ss[i][7],
                    ss[i][2],
                    ss[i][7],
                    ss[i][11],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteA10D(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for i:=0;i<sz;i++{    
        data2:=ss[i][2]
        /*a1:=len(data)
        a2:=len(data2)
        fmt.Println(a1,a2)*/
        if "A10D"==data2{
             w.Write([]string{strb,
                ss[i][2],
                ss[i][7],
                ss[i][11],
            strb+"MAPPING"})
        }
    }
}
func WriteA10C_D(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for i:=0;i<sz;i++{    
        data2:=ss[i][2]
        /*a1:=len(data)
        a2:=len(data2)
        fmt.Println(a1,a2)*/
        if "A10C"==data2||"A10D"==data2{
             w.Write([]string{strb,
                ss[i][2],
                ss[i][7],
                ss[i][11],
            strb+"MAPPING"})
        }
    }
}
func RepeatWriteD(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][3])<len(data)-1{
                continue
            }
            data2:=ss[i][3][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{ss[i][5],
                    ss[i][3],
                    ss[i][5],
                    ss[i][7],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteD(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][3])<len(data)-1{
                continue
            }
            data2:=ss[i][3][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{strb,
                    ss[i][3],
                    ss[i][5],
                    ss[i][7],
                strb+"MAPPING"})
            }
         }
    }
}
func RepeatWriteB(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][9])<len(data)-1{
                continue
            }
            data2:=ss[i][9][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{ss[i][7],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func RepeatWriteZ(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][9])<len(data)-1{
                continue
            }
            data2:=ss[i][9][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{strb,
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteZ(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if len(data)==0{
            break
        }
        for i:=0;i<sz;i++{
            if len(ss[i][9])<len(data)-1{
                continue
            }
            data2:=ss[i][9][0:len(data)-1]+"\n"
            /*a1:=len(data)
            a2:=len(data2)
            fmt.Println(a1,a2)*/
            if data==data2{
                 w.Write([]string{ss[i][9],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func RepeatWriteJ(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if data!="MIRTAZAPINE"+"\n"{
            for i:=0;i<sz;i++{
            data2:=ss[i][5]+"\n"
            if data==data2{
                 w.Write([]string{ss[i][9],
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING"})
            }
         }
        }else{
            for i:=0;i<sz;i++{
                data2:=ss[i][7]+"\n"
                if data==data2{
                     w.Write([]string{ss[i][9],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                    strb+"MAPPING"})
                }
        }
        
    }
}
}
func RepeatWriteJ2(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            if len(ss[i][9])<len(data)-2{
                continue
            }
            data2:=ss[i][9]+"\n"
            if data==data2{
                 w.Write([]string{strb+"MKT",
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteJ(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            if len(data)==0{
                break
            }
            if len(ss[i][9])<len(data)-2{
                continue
            }
            data2:=ss[i][9][0:len(data)-1]+"\n"
            if data==data2{
                 w.Write([]string{ss[i][9],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteAnti(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        if data!="MIRTAZAPINE"+"\n"{
            for i:=0;i<sz;i++{
            data2:=ss[i][5]+"\n"
            if data==data2{
                 w.Write([]string{ss[i][5],
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING"})
            }
         }
        }else{
            for i:=0;i<sz;i++{
                data2:=ss[i][7]+"\n"
                if data==data2{
                     w.Write([]string{strb,
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                    strb+"MAPPING"})
                }
        }
        
    }
}
}
func RepeatWriteFirstOnoT(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][0]+"\n"
            if data==data2{
                 w.Write([]string{ss[i][9],
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteFirstOnoT(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][0]+"\n"
            if data==data2{
                 w.Write([]string{strb,
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}
func WriteFirstOnoE(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][7]+"\n"
            if data==data2{
                 w.Write([]string{strb,
                ss[i][7],
                ss[i][9],
                ss[i][13],
                strb+"MAPPING"})
            }
         }
    }
}

func RepeatFirstE(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        data=strings.ToUpper(data)
        for i:=0;i<sz;i++{
            data2:=ss[i][7]+"\n"
            if data==data2{
                 w.Write([]string{ss[i][9],
                    ss[i][7],
                    ss[i][9],
                    ss[i][13],
                strb+"MAPPING"})
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
    rows:=xlsxRead1.GetRows(sheet)
    lenrows=len(rows)
    s := make([]string,lenrows+2)
    for i:=2;i<=lenrows;i++{
        Place:="A"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue(sheet, Place)
        if len(cell)>0{
            s[i-2]=cell
        }
    }
    readtable=s
}


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
    rows := xlsxRead1.GetRows(sheet)
    lenrows=len(rows)
    for i:=1;i<=lenrows;i++{
        Place:="D"+strconv.Itoa(i)
        cell := xlsxRead1.GetCellValue(sheet, Place)
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


func WriteFirst(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
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
                 w.Write([]string{strb,
                ss[i][1],
                ss[i][3],
                ss[i][5],
                strb+"MAPPING"})
            }
         }
    }
}

func RepeatFirst(str string,strb string,datas[]string){
    addr2="CHPA市场匹配表/数据源/"+str+".CSV"
    xlsxRead2, err := ioutil.ReadFile(addr2)
    if err != nil {
            panic(err)
    }
    r2 := csv.NewReader(strings.NewReader(string(xlsxRead2)))
    ss,_= r2.ReadAll()
    sz = len(ss)
    for _,data:=range datas{
        for i:=0;i<sz;i++{
            data2:=ss[i][1]+"\n"
            if data==data2{
                 w.Write([]string{data,
                ss[i][1],
                ss[i][3],
                ss[i][5],
                strb+"MAPPING"})
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
        cell := xlsxRead1.GetCellValue(sheet, Place)       
        if cell=="Category"{
            i++
            return i 
        }
    }
    return 0
}
func WriteCompsDesc(data1 string,data2 string){
    for i:=0;i<sz;i++{
        if  data1==ss[i][3] {     
             w.Write([]string{data1,
            ss[i][1],
            ss[i][3],
            ss[i][5],
            "Cymbalta CMP _CHPA_ BRAND_1_MAPPING"})
        }
    } 
}
func WriteCompsDesc2(data1 string,data2 string){
    for i:=0;i<sz;i++{
        data:=ss[i][1]
        if data==data2{  

             w.Write([]string{data1,
            ss[i][1],
            ss[i][3],
            ss[i][5],
            "Cymbalta CMP _CHPA_ BRAND_1_MAPPING"})
        }
    } 
}