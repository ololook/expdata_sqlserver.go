package main
import (
        "database/sql"
        _ "github.com/denisenkom/go-mssqldb"
        "log"
        "github.com/tealeg/xlsx"
        "flag"
)
var strFlag = flag.String("sql","","")
var filFlag = flag.String("file","","")
func main() {
        flag.Parse()
        db, err := sql.Open("mssql","server=hostip;port=port;user id=zhangyx;password=zhangyx;database=zhangyx")
        if err != nil {
                log.Fatalf("Open database error: %s\n", err)
        }
        defer db.Close()

        err = db.Ping()
        if err != nil {
                log.Fatal(err)
        }

        rows, err := db.Query(*strFlag)
        if err != nil {
                log.Println(err)
        }
        defer rows.Close()

        var file *xlsx.File
        var sheet *xlsx.Sheet
        var row *xlsx.Row
        var cell *xlsx.Cell
        //var err error

        file = xlsx.NewFile()
        sheet, err = file.AddSheet("Sheet1")
        if err != nil {
                log.Printf(err.Error())
        } 
        cols, _ := rows.Columns()
        row = sheet.AddRow()
        for i:=0;i<len(cols);i++ {
                cell = row.AddCell()
                cell.Value = cols[i]
        }   
        buff := make([]interface{}, len(cols)) 
        data := make([]string, len(cols))  
        for i, _ := range buff {
             buff[i] = &data[i]  
        }
        for rows.Next() {
                row = sheet.AddRow()
                rows.Scan(buff...)
                
                for _, col := range data {
                   cell = row.AddCell()
                   cell.Value = col
                }
                if err != nil {
                        log.Fatal(err)
                }


        }

        err = rows.Err()
        if err != nil {
                log.Fatal(err)
        }

        err = file.Save(*filFlag + ".xlsx")
        if err != nil {
                log.Printf(err.Error())
        }
}

