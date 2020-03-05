namespace CRSource.OnlineTools

open Microsoft.AspNetCore.Http
open System
open System.Linq
open System.Collections.Generic
open System.Text.RegularExpressions
open OfficeOpenXml

module TraceConverter =
    type TraceBlock () =
        member val Position = "" with get, set
        member val Command = "" with get, set
        member val Argument = "" with get, set
        member val Registers = new Dictionary<string, string> () with get, set
        member val Flags = new Dictionary<string, string> () with get, set
        member val Stack = "" with get, set
        member this.TryGetRegister (data : string) =
            match this.Registers.TryGetValue data with
                | true, v ->
                    try Int32.Parse(v, Globalization.NumberStyles.AllowHexSpecifier)
                    with | exn -> 0
                | false, _ ->
                    try Int32.Parse(data, Globalization.NumberStyles.AllowHexSpecifier)
                    with | exn -> 0
        member this.TryGetAddress : string =
            let parsed = Regex.Match(this.Argument, @"\[.+\]")
            if not parsed.Success then
                ""
            else
                let mutable sum = 0
                let mutable ans = ""
                let clean = parsed.Value.Replace("[", "").Replace("]", "")
                let unsignedMember = Regex.Match(clean, @"^\w+")
                let captures = Regex.Matches(clean, @"[+,-]\w+")
                if unsignedMember.Success then
                    let pnum = this.TryGetRegister unsignedMember.Value
                    sum <- sum + pnum
                for cap in captures do
                    let mutable IsNegative = 1
                    if cap.Value.StartsWith('-') then
                        IsNegative <- -IsNegative
                    let pnum = this.TryGetRegister(cap.Value.TrimStart([|'+'; '-'|]))
                    sum <- sum + pnum*IsNegative
                ans <- String.Format("{0:X}", sum)
                while ans.Length < 4 do ans <- "0"+ans
                ans
            
    type TraceBlockS () =
        member val Id = -1 with get, set
        member val Position = "" with get, set
        member val MachineCode = "" with get, set
        member val Command = "" with get, set
        member val Argument = "" with get, set  
        
module TableHelper =
    let WriteCodeHeaders (ws : ExcelWorksheet) =
        ws.Cells.["A1"].Value <- "1st task headers"
    let WriteTraceHeaders (ws : ExcelWorksheet) =
        ws.Cells.["A1"].Value <- "2nd task headers"
    let WriteCodeLine (ws : ExcelWorksheet) (tbs : TraceConverter.TraceBlockS) (id : int) =
        ws.Cells.["A" + (string) id].Value <- tbs.Position
        ws.Cells.["B" + (string) id].Value <- tbs.MachineCode
        ws.Cells.["C" + (string) id].Value <- tbs.Command+" "+tbs.Argument
    let WriteTraceLine (ws : ExcelWorksheet, tb : TraceConverter.TraceBlock, id : int) =
        if id > 3 then
            ws.Cells.["G"+ (string) (id-1)].Value <- tb.Position
        ws.Cells.["A" + (string) id].Value <- tb.Position
        ws.Cells.["B" + (string) id].Value <- tb.Command
        ws.Cells.["C" + (string) id].Value <- tb.Registers.["AX"]
        ws.Cells.["D" + (string) id].Value <- tb.Registers.["BX"]
        ws.Cells.["E" + (string) id].Value <- tb.Registers.["CX"]
        ws.Cells.["F" + (string) id].Value <- tb.Registers.["DX"]
        ws.Cells.["H" + (string) id].Value <- tb.Flags.["OF"]
        ws.Cells.["I" + (string) id].Value <- tb.Flags.["DF"]
        ws.Cells.["J" + (string) id].Value <- tb.Flags.["IF"]
        ws.Cells.["K" + (string) id].Value <- tb.Flags.["SF"]
        ws.Cells.["L" + (string) id].Value <- tb.Flags.["ZF"]
        ws.Cells.["M" + (string) id].Value <- tb.Flags.["AF"]
        ws.Cells.["N" + (string) id].Value <- tb.Flags.["PF"]
        ws.Cells.["O" + (string) id].Value <- tb.Flags.["CF"]
        ws.Cells.["P" + (string) id].Value <- tb.TryGetAddress
        ws.Cells.["Q" + (string) id].Value <- tb.Stack
    let ReadTraceBlock (block : IEnumerable<string>) =
        let tb = new TraceConverter.TraceBlock()
        let lines = new List<string[]>()
        for line in block do
            lines.Add(Regex.Replace(line, @"\s+", " ").Split(' ', System.StringSplitOptions.RemoveEmptyEntries))
        if lines.ToArray().[0].Length = 8 then
            tb.Argument <- lines.ElementAt(0).[2]
            tb.Registers.Add("AX", lines.ElementAt(0).[3].Replace("AX=", ""))
            tb.Registers.Add("SI", lines.ElementAt(0).[4].Replace("SI=", ""))
            tb.Stack <- lines.ElementAt(0).[7]
        else
            tb.Registers.Add("AX", lines.ElementAt(0).[2].Replace("AX=", ""))
            tb.Registers.Add("SI", lines.ElementAt(0).[3].Replace("SI=", ""))
            tb.Stack <- lines.ElementAt(0).[6]
        tb.Position <- lines.ElementAt(0).[0]
        tb.Command <- lines.ElementAt(0).[1]
        tb.Registers.Add("BX", lines.ElementAt(1).[0].Replace("BX=", ""))
        tb.Registers.Add("DI", lines.ElementAt(1).[1].Replace("DI=", ""))
        tb.Registers.Add("CX", lines.ElementAt(2).[8].Replace("CX=", ""))
        tb.Registers.Add("BP", lines.ElementAt(2).[9].Replace("BP=", ""))
        tb.Registers.Add("DX", lines.ElementAt(3).[8].Replace("DX=", ""))
        tb.Registers.Add("SP", lines.ElementAt(3).[9].Replace("SP=", ""))
        tb.Flags.Add("OF", lines.ElementAt(3).[0])
        tb.Flags.Add("DF", lines.ElementAt(3).[1])
        tb.Flags.Add("IF", lines.ElementAt(3).[2])
        tb.Flags.Add("SF", lines.ElementAt(3).[3])
        tb.Flags.Add("ZF", lines.ElementAt(3).[4])
        tb.Flags.Add("AF", lines.ElementAt(3).[5])
        tb.Flags.Add("PF", lines.ElementAt(3).[6])
        tb.Flags.Add("CF", lines.ElementAt(3).[7])
        tb
    let ReadCodeLine (line : string) =
        let tbs = new TraceConverter.TraceBlockS()
        let lines = Regex.Replace(line, @"\s+", " ").Split(' ', System.StringSplitOptions.RemoveEmptyEntries)
        tbs.Position <- Regex.Match(lines.[0], @":\w+").Value.Replace(":", "")
        tbs.MachineCode <- lines.[1]
        tbs.Command <- lines.[2]
        if lines.Length = 4 then
            tbs.Argument <- lines.[3]
        tbs.Id <- try Int32.Parse(tbs.Position, Globalization.NumberStyles.AllowHexSpecifier) with | _ -> -1
        tbs
module ConverterHelper =
    let ConvertToTrace (raw : List<string>) : byte array = 
        let transformed = raw.Skip(3).TakeWhile(fun line -> not (line.Contains "*** End of TRACE buffer ***"))
        let traceBlocks = new List<IEnumerable<string>>()
        for i in [0..((transformed.Count()/4)-1)] do
            traceBlocks.Add(transformed.Skip(4*i).Take(4))
        use ep = new ExcelPackage() in (
            let ws = ep.Workbook.Worksheets.Add("TraceTable")
            let mutable tCounter = 2
            TableHelper.WriteTraceHeaders ws;
            for b in traceBlocks do 
                tCounter <- tCounter + 1
                let tb = TableHelper.ReadTraceBlock(b)
                TableHelper.WriteTraceLine(ws, tb, tCounter)
            let prevVal =
                try Int32.Parse(ws.Cells.["A"+ (string) tCounter].Value.ToString(), Globalization.NumberStyles.AllowHexSpecifier) with
                | _ -> 0
            let mutable lastPos = String.Format("{0:X}", prevVal+1)
            while lastPos.Length < 4 do lastPos <- "0" + lastPos
            ws.Cells.["G" + (string) tCounter].Value <- lastPos
            ep.GetAsByteArray()
        )
    let ConvertToCode (raw : List<string>) : byte array =
        let transofrmed = raw.Skip(3).TakeWhile(fun line -> not (line.Contains "*** End of TRACE buffer ***"))
        use ep = new ExcelPackage() in (
            let ws = ep.Workbook.Worksheets.Add("CodeTable")
            let ul = new SortedList<int, TraceConverter.TraceBlockS>()
            TableHelper.WriteCodeHeaders ws
            for line in transofrmed do
                let tbs = TableHelper.ReadCodeLine(line)
                ul.[tbs.Id] <- tbs
            let mutable tc = 1
            for kvp in ul do
                tc <- tc + 1
                TableHelper.WriteCodeLine ws kvp.Value tc
            ep.GetAsByteArray()
        )