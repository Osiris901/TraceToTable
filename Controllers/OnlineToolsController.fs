namespace CodeRiot.Controllers

open System
open System.Collections.Generic
open System.IO
open System.Text.RegularExpressions

open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Http    

open CRSource.OnlineTools

type OnlineToolsController () =
    inherit Controller()
    
    
    member this.APConverter () =
        this.View()
    
    [<HttpPost>]
    member this.AfdProConverter (upFile : IFormFile) =
        let maxFileSize = Convert.ToInt64 1000000        
        if upFile = null || not this.ModelState.IsValid then
            this.ModelState.AddModelError("При обработке запроса возникла ошибка!", "При загрузке файла возникла ошибка. Попробуйте ещё раз или свяжитесь с разработчиком.")
            this.TempData.["ErrorMsg"] <- "Ошибка при загрузке файла."
        if upFile.Length > maxFileSize then
            this.ModelState.AddModelError("При обработке запроса возникла ошибка!", "Превышен максимальный размер файла.")
            this.TempData.["ErrorMsg"] <- "Превышен максимальный размер файла."
        let fileName = Path.GetExtension upFile.FileName
        if not (fileName = ".txt") && not (fileName = ".TXT") then
            this.ModelState.AddModelError("При обработке запроса возникла ошибка!", "Не поддерживаемый формат файла.")
            this.TempData.["ErrorMsg"] <- "Не поддерживаемый формат файла - (" + fileName + ")"
        if this.TryValidateModel(upFile) then
            try
                let raw = new List<string>()
                use streamReader = new StreamReader(upFile.OpenReadStream()) in
                    while not streamReader.EndOfStream do
                        raw.Add(streamReader.ReadLine())
                if raw.[3].Contains("Stack+0") then
                    let fileContent = ConverterHelper.ConvertToTrace raw
                    let nfn:string = Regex.Replace(upFile.FileName, @"\..+", "") + ".xlsx"
                    this.File(fileContent, "application/octet-stream", nfn) :> ActionResult
                else 
                    let fileContent = ConverterHelper.ConvertToCode raw
                    let nfn:string = Regex.Replace(upFile.FileName, @"\..+", "") + ".xlsx"
                    this.File(fileContent, "application/octet-stream", nfn) :> ActionResult
            with
                | _ ->
                    this.TempData.["ErrorMsg"] <- "Похоже, в вашем файле что-то не так. Убедитесь что он не отредактирован.
                    Если вы уверены, что это наша ошибка - свяжитесь с разработчиком"
                    this.View("APConverter") :> ActionResult
        else
            this.View("APConverter") :> ActionResult