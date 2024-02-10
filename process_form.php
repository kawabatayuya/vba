<?php
if ($_SERVER["REQUEST_METHOD"] == "POST")
{
    $data = $_POST["dataInput"];
    
    // Excelファイルにデータを書き込む処理
    $excelFilePath = "D:\asada_VBA\file.xlsx";
    $excelApp = new COM("Excel.Application") or die("Unable to instantiate Excel");
    $excelApp->Visible = 0;
    $workbook = $excelApp->Workbooks->Open($excelFilePath);
    $worksheet = $workbook->Worksheets(1);
  
    // データをA列に書き込む例
    $lastRow = $worksheet->Cells($worksheet->Rows->Count, "A")->End(-4162)->Row + 1; // -4162 はxlUp
    $worksheet->Cells($lastRow, 1)->Value = $data;
  
    // ファイルを保存して閉じる
    $workbook->Save();
    $workbook->Close(false);
    $excelApp->Quit();
    $excelApp = null;
}
?>