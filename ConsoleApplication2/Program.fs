// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

open System
open System.Configuration
open FSharp.Data
open System.Data
open System.Data.Linq

open FSharp.Data.CsvExtensions

open FSharp.Charting
open FSharp.Charting.ChartTypes

open System.Drawing
open System.Windows.Forms

open MySql.Data.MySqlClient

// Excel操作用
open Microsoft.Office.Interop


type Stocks = CsvProvider<"http://ichart.finance.yahoo.com/table.csv?s=MSFT">

[<EntryPoint>]
let main argv =


    let msft = Stocks.Load("http://ichart.finance.yahoo.com/table.csv?s=MSFT")

    // 最新の行をチェックする。なお 'Date' プロパティは
    // 'DateTime' 型で、 'Open' プロパティは 'decimal' 型であることに注意
    let firstRow = msft.Rows |> Seq.head
    let lastDate = firstRow.Date
    let lastOpen = firstRow.Open


    //for row in msft.Rows do
        // printfn "HLOC: (%s, %s, %s)" (row.GetColumn "High") (row.GetColumn "Low") (row.GetColumn "Date")
        //printfn "HLOC: (%f, %M, %O)" (row.["High"].AsFloat()) (row?Low.AsDecimal()) (row?Date.AsDateTime())

    // 終値が始値よりも高いもののうち、上位10位の株価をTSV(タブ区切り)形式で保存します：
    // msft.Filter(fun row -> row?Close.AsFloat() > row?Open.AsFloat())
    //    .Truncate(10)
    //    .SaveToString('\t')

    //for row in msft.Rows do
        // printfn "HLOC: (%s, %s, %s)" (row.GetColumn "High") (row.GetColumn "Low") (row.GetColumn "Date")
    //    printfn "HLOC: (%A, %A, %A, %A)" row.High row.Low row.Open row.Close

  
    let mychart = [ for row in msft.Rows -> row.Date,row.Open ]
                    |> Chart.FastLine 
    let myChartControl = new ChartControl(mychart,Dock=DockStyle.Fill)
    let lbl = new Label(Text="my label")
    let form = new Form(Visible=true,TopMost=true,Width=700,Height=500)
    form.Controls.Add lbl
    form.Controls.Add(myChartControl)
    do Application.Run(form) |> ignore

    let conn = new MySqlConnection "Server=localhost;Database=myfsharp_db;Uid=kei;Pwd=8wkmrlv;"
    conn.Open()

    let cmd = new MySqlCommand("select * from iris",conn)
    let reader = cmd.ExecuteReader()

    // while reader.Read()
    //   do System.Console.WriteLine(reader.GetString 0)

    let xlApp = new Excel.ApplicationClass()
    let xlWorkBook = xlApp.Workbooks.Add()
    let xlWorkSheet = xlWorkBook.Worksheets.[1] :?> Excel.Worksheet

   
    let retVal = ([| while(reader.Read())
                            do yield(reader.GetString("sep_len"),reader.GetString("sep_wid")) |])

    

    // xlWorkSheet.Range("A1","E30").Value2 <- retVal :?> Excel.Range


    for i in 1 .. retVal.Length-1 do
        let (var1,var2) = retVal.[i]
        xlWorkSheet.Cells.[i,1] <- var1
        xlWorkSheet.Cells.[i,2] <- var2

    System.Console.WriteLine(retVal)

    xlApp.Visible <- true

    0 // return an integer exit code
